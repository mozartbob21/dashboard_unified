"""Private tool jobs and shared organization presentation assets."""
import asyncio
import importlib.util
import json
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from core.roles import check_module_access, is_full_access
from core.web import templates
from services.tools import branding, coordinates, documents, excel_merge, exe_builder, macros, images, external_images, pdf_pages, pptx_converter as pptx
from services.tools.workspace import (ToolError, account_root, job_path, read_upload,
                                      safe_name, start_job, MAX_FILE, MAX_TOTAL)


def require_tools(request: Request):
    user = getattr(request.state, 'user', None)
    if not user or not check_module_access(user, 'tools'):
        raise HTTPException(403, 'Нет доступа к инструментам')


router = APIRouter(prefix='/tools', dependencies=[Depends(require_tools)])


def fail(message):
    return JSONResponse({'ok':False, 'message':str(message)}, status_code=400)


@router.get('')
async def page(request: Request):
    return templates.TemplateResponse(request, 'tools.html', {'request':request,
        'user_username':request.state.user.get('username',''), 'can_build':is_full_access(request.state.user)})


@router.get('/api/options')
async def options(request: Request):
    return {**await asyncio.to_thread(branding.options), 'pdf_docx':bool(importlib.util.find_spec('pdf2docx')),
            'exe':exe_builder.available(), 'can_build':is_full_access(request.state.user)}


@router.post('/api/pdf-info')
async def pdf_info(request: Request):
    try:
        form = await request.form(max_files=20, max_fields=1, max_part_size=MAX_FILE)
        files = [await read_upload(upload, {'.pdf'}) for upload in form.getlist('file')]
        if sum(len(data) for _, data in files) > MAX_TOTAL:
            raise ToolError('Общий размер файлов превышает 60 МБ.')
        return {'files': await asyncio.to_thread(pdf_pages.info, files)}
    except ToolError as exc:
        return fail(exc)


@router.post('/api/run')
async def run(request: Request):
    try:
        form=await request.form(max_files=40, max_fields=15, max_part_size=MAX_FILE)
        kind=str(form.get('tool') or '')
        allowed={'html_pptx':{'.html','.htm'}, 'pdf_docx':{'.pdf'}, 'pdf_pages':{'.pdf'},
                 'merge':{'.xlsx','.xls','.xlsb'}, 'coordinates':{'.xlsx'}, 'exe':{'.py'}, 'macros':set()}
        if kind not in allowed:
            raise ToolError('Выберите инструмент.')
        if kind=='exe' and not is_full_access(request.state.user):
            raise HTTPException(403,'Сборка EXE доступна только администратору.')
        files=[]
        for upload in form.getlist('file'):
            files.append(await read_upload(upload,allowed[kind]))
            if sum(len(data) for _,data in files)>MAX_TOTAL:
                raise ToolError('Общий размер файлов превышает 60 МБ.')
        if kind not in ('html_pptx','macros','merge','pdf_pages') and len(files)!=1:
            raise ToolError('Выберите один файл.')
        if kind=='merge' and not files:
            raise ToolError('Выберите Excel-файлы или папку.')
        pictures=[]
        if kind=='html_pptx':
            for upload in form.getlist('images'):
                if getattr(upload,'filename',''):
                    pictures.append(await read_upload(upload,{'.png','.jpg','.jpeg','.webp'}))
            if sum(len(data) for _,data in files+pictures)>MAX_TOTAL:
                raise ToolError('Общий размер файлов превышает 60 МБ.')
        root=account_root(request.state.user)
        if kind=='macros':
            prompt=str(form.get('prompt') or '');target=str(form.get('target') or 'Excel VBA')
            operation=lambda job:macros.generate(prompt,target,job)
        elif kind=='merge':
            operation=lambda job:excel_merge.merge(files,job)
        elif kind=='coordinates':
            imported = None
            upload = form.get('registry')
            if getattr(upload, 'filename', ''):
                _, imported = await read_upload(upload, {'.xlsx'})
            radius = str(form.get('radius') or '500')
            prefix = str(form.get('prefix') or '')
            operation=lambda job:coordinates.group(*files[0],radius,prefix,root,job,imported)
        elif kind=='exe':
            operation=lambda job:exe_builder.build(*files[0],form.get('windowed')=='true',job)
        elif kind=='pdf_docx':
            operation=lambda job:documents.convert(kind,*files[0],job)
        elif kind=='pdf_pages':
            action = str(form.get('action') or '')
            pages = str(form.get('pages') or '')
            angle = str(form.get('angle') or '90')
            operation=lambda job:pdf_pages.process(files,action,pages,angle,job)
        else:
            if len(files)>1:
                raise ToolError('Выберите один HTML-файл.')
            html=files[0][1].decode('utf-8-sig',errors='replace') if files else str(form.get('html') or '')
            if not html.strip() or len(html.encode('utf-8'))>MAX_FILE:
                raise ToolError('Укажите HTML размером до 20 МБ.')
            mode=str(form.get('mode') or 'smart')
            if mode not in ('smart','ai','shots'):
                raise ToolError('Неизвестный режим конвертации.')
            name=str(form.get('template') or '')
            assets = await asyncio.to_thread(branding.capture, name if mode != 'shots' else '')
            load_remote=form.get('remote_images')=='true'
            def operation(job):
                with pptx.workspace(branding.prepare_job(job, assets)):
                    source_html=images.embed_uploads(html,pictures)
                    missing=0
                    if load_remote:source_html,missing=external_images.embed_remote(source_html)
                    output=str((job/'presentation.pptx').resolve())
                    if mode=='shots':
                        shots=pptx.render_slides_images(source_html,str((job/'input').resolve()))
                        if not shots:
                            raise ToolError('Не удалось получить слайды из HTML.')
                        pptx.build_pptx_from_images(shots,output)
                        import shutil
                        shutil.rmtree(str((job/'input').resolve())+'_shots',ignore_errors=True)
                        count=len(shots)
                    else:
                        try:
                            slides=pptx.parse_html_to_slides_ai(source_html) if mode=='ai' else pptx.parse_html_to_slides_pro(source_html)
                        except Exception:
                            raise ToolError('ГосЧат недоступен. Повторите или выберите режим «Структура HTML».')
                        if not slides:
                            raise ToolError('Слайды не распознаны. Попробуйте режим «Структура HTML».')
                        pptx.build_pptx(slides,name or None,output)
                        count=len(slides)
                    return {'file':'presentation.pptx','message':f'Презентация готова. Слайдов: {count}.' + (f' Не загрузилось картинок по ссылкам: {missing}. Их можно добавить с компьютера.' if missing else '')}
        job_id=start_job(root,kind,operation)
        return {'ok':True,'job_id':job_id}
    except ToolError as exc:
        return fail(exc)


@router.get('/api/jobs/{job_id}')
async def status(job_id: str, request: Request):
    try:
        path=job_path(account_root(request.state.user),job_id)/'status.json'
        state=json.loads(path.read_text(encoding='utf-8'))
    except (ToolError,OSError,json.JSONDecodeError):
        raise HTTPException(404,'Задание не найдено')
    if state.get('state')=='done' and state.get('file'):
        from urllib.parse import quote
        state['url']=f'/tools/download/{job_id}/'+quote(state['file'])
    return state


@router.get('/download/{job_id}/{name}')
async def download(job_id: str, name: str, request: Request):
    try:
        name=safe_name(name)
        root=job_path(account_root(request.state.user),job_id)
        state=json.loads((root/'status.json').read_text(encoding='utf-8'))
        path=root/name
        if state.get('state')!='done' or state.get('file')!=name or not path.is_file() or path.is_symlink():
            raise ValueError()
    except (ToolError,ValueError,OSError):
        raise HTTPException(404,'Файл не найден')
    return FileResponse(path,filename=name,headers={'Cache-Control':'no-store'})


@router.post('/html2pptx/preview')
async def preview(request: Request):
    form=await request.form(max_files=1,max_fields=3,max_part_size=MAX_FILE)
    html=str(form.get('html') or '')
    if not html.strip() or len(html)>MAX_FILE:
        return fail('Укажите HTML размером до 20 МБ.')
    return {'ok':True,'slides':pptx.parse_html_to_slides_pro(html)}


@router.post('/template')
async def upload_template(request: Request):
    try:
        form=await request.form(max_files=1,max_fields=2)
        name,data=await read_upload(form.get('template_file'),{'.pptx'})
        return {'ok':True, **await asyncio.to_thread(branding.save_template, name, data)}
    except ToolError as exc:
        return fail(exc)


@router.post('/template/delete')
async def delete_template(request: Request):
    try:
        form=await request.form(max_files=0,max_fields=2)
        name=safe_name(form.get('name'),{'.pptx'})
        return {'ok':True, **await asyncio.to_thread(branding.delete_template, name)}
    except ToolError as exc:
        return fail(exc)


@router.post('/emblem')
async def upload_emblem(request: Request):
    try:
        form=await request.form(max_files=1,max_fields=2)
        name,data=await read_upload(form.get('emblem_file'),{'.png','.jpg','.jpeg'})
        return {'ok':True, **await asyncio.to_thread(branding.save_emblem, data)}
    except (ToolError,ValueError,OSError):
        return fail('Не удалось прочитать изображение (PNG/JPEG до 16 мегапикселей).')


@router.get('/emblem')
async def emblem():
    path = (await asyncio.to_thread(branding.root)) / 'emblem.png'
    if not path.is_file() or path.is_symlink():
        raise HTTPException(404, 'Общий герб ещё не загружен')
    return FileResponse(path, media_type='image/png', headers={'Cache-Control': 'no-store'})
