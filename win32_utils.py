import shutil
import sys
import os
import time

def clear_gen_py():
    """
    Attempts to clear the win32com gen_py cache folder.
    This resolves the "AttributeError: module 'win32com.gen_py...' has no attribute 'CLSIDToPackageMap'" error.
    """
    try:
        import win32com
        if hasattr(win32com, "__gen_path__"):
            gen_path = win32com.__gen_path__
            if os.path.isdir(gen_path):
                print(f"[win32_utils] Clearing corrupted cache at: {gen_path}")
                shutil.rmtree(gen_path)
                return True
    except Exception:
        pass
    
    # Fallback: check standard %LOCALAPPDATA%\Temp\gen_py
    local_app_data = os.environ.get('LOCALAPPDATA')
    if local_app_data:
        temp_gen_py = os.path.join(local_app_data, 'Temp', 'gen_py')
        if os.path.isdir(temp_gen_py):
            print(f"[win32_utils] Clearing corrupted cache at: {temp_gen_py}")
            try:
                shutil.rmtree(temp_gen_py)
                return True
            except Exception as e:
                print(f"[win32_utils] Failed to delete {temp_gen_py}: {e}")
    
    return False

def _is_gen_py_error(exc: Exception) -> bool:
    msg = str(exc)
    return "CLSIDToPackageMap" in msg or "gen_py" in msg

def safe_dispatch(prog_id, retries: int = 2, delay_sec: float = 1.0, use_dispatch_ex: bool = True):
    """
    Wrapper for win32com.client.Dispatch that handles the common
    gen_py cache corruption error by clearing the cache and retrying.
    """
    import win32com.client

    last_exc = None
    for attempt in range(retries + 1):
        try:
            if use_dispatch_ex and hasattr(win32com.client, "DispatchEx"):
                return win32com.client.DispatchEx(prog_id)
            return win32com.client.Dispatch(prog_id)
        except AttributeError as e:
            last_exc = e
            if _is_gen_py_error(e):
                print(f"[win32_utils] Encountered win32com cache corruption: {e}")
                if clear_gen_py():
                    print("[win32_utils] Cache cleared. Retrying Dispatch...")
                    time.sleep(delay_sec)
                    continue
            raise
        except Exception as e:
            last_exc = e
            if attempt < retries:
                time.sleep(delay_sec)
                continue
            raise

    if last_exc:
        raise last_exc
    raise RuntimeError("safe_dispatch failed without exception.")
