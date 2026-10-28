"""Native rating module; all data, jobs and exports require a current grant."""
import json
from fastapi import APIRouter,BackgroundTasks,Depends,HTTPException,Request
from fastapi.responses import JSONResponse,Response
from core.web import templates
from services.omsu_rating import store
from services.omsu_rating import export as rating_export

PRIVATE={'Cache-Control':'private, no-store'}

def require_rating(request:Request):
    user=getattr(request.state,'user',None)
    if not user:
        from services.auth.security import get_user_from_token
        user=get_user_from_token(request.cookies.get('access_token'))
    if not user:raise HTTPException(401,'Войдите в Нейрону.')
    if user.get('kc_sub'):
        from core.roles import check_module_access
        allowed=check_module_access(user,'omsu-rating')
    else:
        from utils.db import get_db_connection
        if type(user.get('id')) is not int or user['id']<=0:raise HTTPException(401,'Войдите в Нейрону.')
        with get_db_connection() as conn:
            row=conn.execute('SELECT is_active,modules FROM users WHERE id=? AND NOT EXISTS (SELECT 1 FROM account_archives WHERE user_id=users.id)',(user['id'],)).fetchone()
        try:
            modules=json.loads(row['modules'] or '[]') if row else []
            allowed=bool(row and row['is_active'] and isinstance(modules,list) and 'omsu-rating' in modules)
        except (ValueError,TypeError):allowed=False
    if not allowed:raise HTTPException(403,'Нет доступа к рейтингу ОМСУ.')
    return user

router=APIRouter(prefix='/omsu-rating',dependencies=[Depends(require_rating)])

@router.get('')
def page(request:Request):return templates.TemplateResponse(request,'omsu_rating.html',{'request':request},headers=PRIVATE)

@router.get('/api/data')
def data():return JSONResponse(store.read(),headers=PRIVATE)

@router.get('/api/status')
def status():return JSONResponse(store.status(),headers=PRIVATE)

@router.post('/api/refresh')
def refresh(request:Request,background:BackgroundTasks):
    origin=request.headers.get('origin','')
    expected=f'{request.url.scheme}://{request.url.netloc}'
    if origin!=expected or request.headers.get('sec-fetch-site')=='cross-site':raise HTTPException(403,'Обновление доступно только со страницы Нейроны.')
    token=store.claim()
    if token is not None:background.add_task(store.refresh,token)
    return JSONResponse({'status':'running','already_running':token is None},status_code=202,headers=PRIVATE)

@router.get('/export.xlsx')
def export(ind:int|None=None,period:int|None=None):
    if ind is not None and ind not in (4,5,6,7):raise HTTPException(400,'Неизвестный показатель.')
    if period is not None and period not in range(4):raise HTTPException(400,'Неизвестный период.')
    payload=rating_export.workbook_bytes(store.read(),ind,period)
    return Response(payload,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={**PRIVATE,'Content-Disposition':'attachment; filename="omsu-rating.xlsx"'})


@router.get('/export.html')
def snapshot():
    """Portable interactive snapshot; its CSP blocks every network request."""
    return Response(rating_export.snapshot_html(store.read()),media_type='text/html',headers={**PRIVATE,'Content-Disposition':'attachment; filename="omsu-rating-snapshot.html"','X-Content-Type-Options':'nosniff'})
