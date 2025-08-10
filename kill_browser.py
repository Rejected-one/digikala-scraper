import psutil

def kill_process(process_name: str) -> None:
    """
    بستن تمام فرایندهایی با نام مشخص‌شده (مثلاً: 'msedge.exe')
    """
    killed_any = False

    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info.get('name', '').lower() == process_name.lower():
                proc.kill()
                print(f"✅ {process_name} (PID {proc.pid}) بسته شد.")
                killed_any = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            print(f"⚠️ خطا در بستن PID {proc.pid} ({process_name})")

    if not killed_any:
        print(f"❌ هیچ فرایندی با نام «{process_name}» پیدا نشد.")
