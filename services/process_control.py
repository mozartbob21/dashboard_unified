"""Contain descendants of a Windows collector; POSIX uses a process group."""
import os
import threading

_active = set()
_active_lock = threading.Lock()


def set_active(module, active):
    with _active_lock:
        if active:
            _active.add(module)
        else:
            _active.discard(module)


def is_active(module):
    with _active_lock:
        return module in _active


def bounded_lines(stream, expired):
    """A descendant retaining a pipe must not defeat the task deadline."""
    import queue
    import threading
    values = queue.Queue(maxsize=128)
    done = threading.Event()
    def read():
        try:
            for line in stream:
                while not done.is_set() and not expired.is_set():
                    try:
                        values.put(line, timeout=.2)
                        break
                    except queue.Full:
                        pass
                if done.is_set() or expired.is_set():
                    break
        finally:
            done.set()
    threading.Thread(target=read, daemon=True, name='collector-output').start()
    try:
        while not expired.is_set():
            try:
                yield values.get(timeout=.2)
            except queue.Empty:
                if done.is_set():
                    break
    finally:
        done.set()


def windows_job(process):
    if os.name != 'nt':
        return None
    import ctypes
    from ctypes import wintypes as w
    class Basic(ctypes.Structure):
        _fields_ = [('PerProcessUserTimeLimit', ctypes.c_longlong), ('PerJobUserTimeLimit', ctypes.c_longlong),
                    ('LimitFlags', w.DWORD), ('MinimumWorkingSetSize', ctypes.c_size_t),
                    ('MaximumWorkingSetSize', ctypes.c_size_t), ('ActiveProcessLimit', w.DWORD),
                    ('Affinity', ctypes.c_size_t), ('PriorityClass', w.DWORD), ('SchedulingClass', w.DWORD)]
    class IO(ctypes.Structure):
        _fields_ = [(name, ctypes.c_ulonglong) for name in ('ReadOperationCount', 'WriteOperationCount',
                    'OtherOperationCount', 'ReadTransferCount', 'WriteTransferCount', 'OtherTransferCount')]
    class Extended(ctypes.Structure):
        _fields_ = [('BasicLimitInformation', Basic), ('IoInfo', IO), ('ProcessMemoryLimit', ctypes.c_size_t),
                    ('JobMemoryLimit', ctypes.c_size_t), ('PeakProcessMemoryUsed', ctypes.c_size_t),
                    ('PeakJobMemoryUsed', ctypes.c_size_t)]
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateJobObjectW.argtypes = [ctypes.c_void_p, w.LPCWSTR]
    kernel.CreateJobObjectW.restype = w.HANDLE
    kernel.SetInformationJobObject.argtypes = [w.HANDLE, ctypes.c_int, ctypes.c_void_p, w.DWORD]
    kernel.AssignProcessToJobObject.argtypes = [w.HANDLE, w.HANDLE]
    kernel.CloseHandle.argtypes = [w.HANDLE]
    handle = kernel.CreateJobObjectW(None, None)
    info = Extended()
    info.BasicLimitInformation.LimitFlags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    if not handle:
        raise OSError('Не удалось создать ограничитель процесса.')
    if (not kernel.SetInformationJobObject(handle, 9, ctypes.byref(info), ctypes.sizeof(info))
            or not kernel.AssignProcessToJobObject(handle, w.HANDLE(int(process._handle)))):
        kernel.CloseHandle(handle)
        raise OSError('Не удалось ограничить дочерние процессы блока.')
    import threading
    lock = threading.Lock()
    def close():
        nonlocal handle
        with lock:
            if handle:
                kernel.CloseHandle(handle)
                handle = None
    return close
