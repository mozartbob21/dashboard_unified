from pathlib import Path
from datetime import date
from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Request
from fastapi.responses import FileResponse
from core.roles import check_module_access
from services.auth.integrations import credentials
from services.edds import runner, arm


def require_edds(request: Request):
    if not check_module_access(getattr(request.state,'user',None),'edds'):
        raise HTTPException(403,'Нет доступа к блоку ЕДДС')


router=APIRouter(prefix='/edds',dependencies=[Depends(require_edds)])


@router.get('')
async def page():
    return FileResponse(Path(__file__).resolve().parents[1]/'services/edds/dashboard.html',media_type='text/html',headers={'Cache-Control':'no-store'})


@router.get('/water-daily')
async def water_daily():
    return runner.water_daily()


@router.get('/status')
async def status():
    return {**runner.status(), 'arm_configured': bool(credentials('edds_arm')),
            'complaints_configured': bool(credentials('edds'))}


@router.get('/arm/report')
def arm_report(from_date: date, to_date: date, coordinates: bool = False):
    # A sync route runs blocking portal requests in FastAPI's worker thread pool.
    from fastapi.responses import JSONResponse
    try:
        return JSONResponse({'grid': arm.fetch_report(from_date, to_date, coordinates)},
                            headers={'Cache-Control': 'no-store'})
    except arm.ArmError as error:
        raise HTTPException(error.status, str(error), headers={'Cache-Control': 'no-store'}) from None


@router.post('/refresh',status_code=202)
async def refresh(request: Request, background: BackgroundTasks):
    if not credentials(): raise HTTPException(400,'Логин и пароль Добродела не настроены. Обратитесь к администратору.')
    if runner.status()['running']: raise HTTPException(409,'Сбор жалоб уже выполняется')
    background.add_task(runner.run,request.state.user.get('username','—'))
    return {'ok':True,'message':'Сбор поставлен в очередь. Обновление может занять несколько минут.'}
