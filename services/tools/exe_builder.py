"""Build inside Windows Sandbox, never run user Python on the Neurona host."""
import ast
import os
import shutil
import subprocess
import threading
import time
from pathlib import Path
from xml.etree import ElementTree as ET
from .workspace import ToolError

LOCK = threading.Lock()


def runtime_dir():
    path = os.getenv('TOOLS_BUILD_RUNTIME', '').strip()
    return Path(path) if path else None


def available():
    runtime = runtime_dir()
    return os.name == 'nt' and bool(shutil.which('WindowsSandbox.exe')) and bool(runtime and (runtime/'python.exe').is_file())


def sandbox_xml(source, output, runtime):
    root = ET.Element('Configuration')
    for key,value in [('VGpu','Disable'),('Networking','Disable'),('ClipboardRedirection','Disable'),
                      ('PrinterRedirection','Disable'),('AudioInput','Disable'),('VideoInput','Disable'),('MemoryInMB','4096')]:
        ET.SubElement(root,key).text=value
    folders=ET.SubElement(root,'MappedFolders')
    for host,guest,readonly in [(source,r'C:\input',True),(output,r'C:\result',False),(runtime,r'C:\runtime',True)]:
        folder=ET.SubElement(folders,'MappedFolder')
        ET.SubElement(folder,'HostFolder').text=str(Path(host).resolve())
        ET.SubElement(folder,'SandboxFolder').text=guest
        ET.SubElement(folder,'ReadOnly').text=str(readonly).lower()
    ET.SubElement(ET.SubElement(root,'LogonCommand'),'Command').text=r'cmd.exe /d /c C:\input\build.cmd'
    return ET.tostring(root,encoding='unicode')


def build(name, data, windowed, job):
    try:
        ast.parse(data, filename='source.py')
    except (SyntaxError, ValueError, UnicodeError):
        raise ToolError('В Python-файле ошибка синтаксиса или кодировки.')
    if not available():
        raise ToolError('Сборка EXE требует Windows Sandbox и отдельного Python с PyInstaller. Настройка описана в docs/tools-and-mingkh.md.')
    if not LOCK.acquire(blocking=False):
        raise ToolError('Другая сборка EXE уже выполняется. Повторите после её завершения.')
    try:
        source=job/'input';source.mkdir()
        output=job/'sandbox-result';output.mkdir()
        (source/'source.py').write_bytes(data)
        option=' --windowed' if windowed else ''
        command='@echo off\r\nmkdir C:\\work\r\nC:\\runtime\\python.exe -I -m PyInstaller --clean --noconfirm --onefile --name application --distpath C:\\result --workpath C:\\work --specpath C:\\work'+option+' C:\\input\\source.py > C:\\result\\build.log 2>&1\r\necho %errorlevel% > C:\\result\\done.txt\r\nshutdown /s /t 0\r\n'
        (source/'build.cmd').write_bytes(command.encode('ascii'))
        config=job/'build.wsb';config.write_text(sandbox_xml(source,output,runtime_dir()),encoding='utf-8')
        process=subprocess.Popen([shutil.which('WindowsSandbox.exe'),str(config)],shell=False)
        done=output/'done.txt';exe=output/'application.exe'
        deadline=time.monotonic()+600
        # New Windows versions can return from the launcher while the guest is running.
        while not done.exists():
            if process.poll() not in (None,0):
                raise ToolError('Windows Sandbox не запустилась. Проверьте её запуск на компьютере-сервере.')
            if time.monotonic()>=deadline:
                if process.poll() is None:process.kill()
                raise ToolError('Сборка превысила 10 минут. Проверьте Windows Sandbox на сервере.')
            time.sleep(1)
        try:process.wait(timeout=30)
        except subprocess.TimeoutExpired:process.kill()
        if not done.exists() or done.read_text().strip()!='0' or not exe.exists():
            raise ToolError('EXE не собран. В изолированном Python должны быть установлены PyInstaller и зависимости скрипта.')
        if exe.stat().st_size > 250*1024**2:
            raise ToolError('Готовый EXE превышает 250 МБ.')
        shutil.move(exe, job/'application.exe')
        return {'file':'application.exe','message':'EXE собран в Windows Sandbox без доступа к сети.'}
    finally:
        shutil.rmtree(job/'sandbox-result',ignore_errors=True)
        (job/'build.wsb').unlink(missing_ok=True)
        LOCK.release()
