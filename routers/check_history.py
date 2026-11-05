"""Read-only, account-scoped views; no model, dispatcher or external I/O."""
from datetime import date
from fastapi import APIRouter, Depends, HTTPException, Query, Request
from fastapi.responses import JSONResponse
from services import check_analysis, check_history as history, check_metrics
from services.auth.security import get_user_from_token

router = APIRouter(prefix='/api/check-history', tags=['check-history'])

def require_account(request: Request):
    user = get_user_from_token(request.cookies.get('access_token'))
    if not user or not user.get('is_active',True):
        raise HTTPException(401,'Требуется авторизация.')
    return user

def filters(operation, source=None):
    if operation is not None and operation not in check_metrics.OPERATIONS:
        raise HTTPException(400,'Неизвестный вид проверки.')
    if source is not None and source not in check_metrics.SOURCE_IDS:
        raise HTTPException(400,'Неизвестный источник.')
    return {k:v for k,v in [('operation',operation),('source',source)] if v}

def response(value):
    return JSONResponse(value,headers={'Cache-Control':'private, no-store'})

@router.get('/{module_id}')
def listing(module_id: str, operation: str|None=None, source: str|None=None, start: date|None=None,
            end: date|None=None, status: str|None=None, limit: int=Query(20,ge=1,le=100), user=Depends(require_account)):
    if start and end and start>end:
        raise HTTPException(400,'Начало периода позже окончания.')
    return response({'checks':history.read_checks(user,module_id,start=start,end=end,status=status,limit=limit,
                                                  params=filters(operation,source))})

@router.get('/{module_id}/analysis')
def analysis(module_id: str, operation: str|None=None, source: str|None=None, window: str='last_two', user=Depends(require_account)):
    result = history.compare(user,module_id,window=window,params=filters(operation,source))
    return response(check_analysis.describe(result,compact=True))
