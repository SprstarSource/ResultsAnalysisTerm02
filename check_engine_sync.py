import sys
import os
import re

# إجبار بيئة ويندوز على استخدام ترميز UTF-8 لطباعة النصوص العربية بدون أخطاء
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

def check_error_handling(file_path):
    if not os.path.exists(file_path):
        print(f"[ERROR] لم يتم العثور على الملف: '{file_path}'")
        sys.exit(1)

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(file_path, 'r', encoding='windows-1256', errors='ignore') as f:
            lines = f.readlines()

    proc_start_re = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property)\s+([a-zA-Z0-9_]+)', re.IGNORECASE)
    proc_end_re = re.compile(r'^\s*End\s+(?:Sub|Function|Property)', re.IGNORECASE)
    error_handler_re = re.compile(r'^\s*On\s+Error\s+(?:GoTo|Resume\s+Next)', re.IGNORECASE)

    current_proc = None
    has_error_handler = False
    missing_handlers = []

    for line in lines:
        start_match = proc_start_re.match(line)
        if start_match:
            current_proc = start_match.group(1)
            has_error_handler = False
            continue

        if current_proc:
            if error_handler_re.search(line):
                has_error_handler = True
            
            if proc_end_re.match(line):
                if not has_error_handler:
                    missing_handlers.append(current_proc)
                current_proc = None

    if missing_handlers:
        print(f"[FAIL] فشل الفحص: تم العثور على دوال بدون معالج أخطاء في الموديول ({os.path.basename(file_path)}):")
        for proc in missing_handlers:
            print(f"  - {proc}")
        sys.exit(1)
    else:
        print(f"[PASS] نجاح الفحص: جميع الدوال في الموديول ({os.path.basename(file_path)}) محمية بمعالج أخطاء.")
        sys.exit(0)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python check_engine_sync.py <module.bas>")
        sys.exit(1)
    check_error_handling(sys.argv[1])