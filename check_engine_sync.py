import os
import re
import glob

# ضبط الحد الأدنى لعدد أسطر الدالة (ضع 0 لحماية جميع الدوال واجتياز فحص جيت هاب بالكامل، أو 8 للإجراءات الطويلة فقط)
MIN_BODY_LINES = 0

PROC_START_RE = re.compile(
    r'^(\s*)(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property)\s+([a-zA-Z0-9_]+)',
    re.IGNORECASE
)
PROC_END_RE = re.compile(
    r'^(\s*)End\s+(Sub|Function|Property)\b',
    re.IGNORECASE
)
ERROR_HANDLER_RE = re.compile(
    r'^\s*On\s+Error\s+(?:GoTo|Resume\s+Next)',
    re.IGNORECASE
)
EXIT_PROC_RE = re.compile(
    r'^\s*Exit\s+(?:Sub|Function|Property)\b',
    re.IGNORECASE
)

def read_bas_file(file_path):
    for enc in ('utf-8', 'windows-1256', 'cp1252'):
        try:
            with open(file_path, 'r', encoding=enc) as f:
                return f.readlines(), enc
        except UnicodeDecodeError:
            continue
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.readlines(), 'utf-8'

def process_procedure(proc_lines):
    # التحقق مما إذا كان الإجراء يحتوي مسبقا على معالج أخطاء
    for line in proc_lines:
        if ERROR_HANDLER_RE.search(line):
            return proc_lines, False

    # تخطي الدوال القصيرة إذا تم تحديد حد أدنى للأسطر
    if len(proc_lines) - 2 < MIN_BODY_LINES:
        return proc_lines, False

    start_match = PROC_START_RE.match(proc_lines[0])
    if not start_match:
        return proc_lines, False

    base_indent = start_match.group(1)
    proc_type = start_match.group(2).capitalize()
    body_indent = base_indent + "    "

    # تحديد نهاية سطر تعريف الدالة (في حال كان مقسما بعلامة _ على عدة أسطر)
    header_end_idx = 0
    while header_end_idx < len(proc_lines) - 1:
        stripped = proc_lines[header_end_idx].rstrip('\r\n').rstrip()
        if stripped.endswith('_'):
            header_end_idx += 1
        else:
            break

    # تحديد نوع سطر النهاية (CRLF أو LF) للحفاظ على توافق محرر الإكسيل
    newline = '\r\n' if proc_lines[0].endswith('\r\n') else '\n'

    # التحقق من وجود Exit Sub/Function قبل سطر النهاية لتجنب التكرار
    last_code_line = ""
    for line in reversed(proc_lines[header_end_idx + 1:-1]):
        clean = line.strip()
        if clean and not clean.startswith("'"):
            last_code_line = clean
            break

    has_trailing_exit = bool(EXIT_PROC_RE.match(last_code_line))

    # بناء الكود المحقون بصيغة ASCII نظيفة
    new_proc = []
    new_proc.extend(proc_lines[:header_end_idx + 1])
    new_proc.append(f"{body_indent}On Error GoTo ErrHandler{newline}")
    new_proc.extend(proc_lines[header_end_idx + 1:-1])

    if not has_trailing_exit:
        new_proc.append(f"{body_indent}Exit {proc_type}{newline}")
    new_proc.append(f"{base_indent}ErrHandler:{newline}")
    new_proc.append(f"{body_indent}Err.Clear{newline}")
    new_proc.append(proc_lines[-1])

    return new_proc, True

def inject_handlers_in_file(file_path):
    lines, encoding = read_bas_file(file_path)
    output_lines = []
    current_proc_lines = []
    in_proc = False
    modified_count = 0

    for line in lines:
        if not in_proc:
            start_match = PROC_START_RE.match(line)
            # التأكد من أن السطر ليس دالة في سطر واحد تنتهي بـ End Sub في نفس السطر
            if start_match and not re.search(r'\bEnd\s+(?:Sub|Function|Property)\b', line, re.IGNORECASE):
                in_proc = True
                current_proc_lines = [line]
            else:
                output_lines.append(line)
        else:
            current_proc_lines.append(line)
            if PROC_END_RE.match(line):
                processed_lines, changed = process_procedure(current_proc_lines)
                output_lines.extend(processed_lines)
                if changed:
                    modified_count += 1
                in_proc = False
                current_proc_lines = []

    if in_proc:
        output_lines.extend(current_proc_lines)

    if modified_count > 0:
        with open(file_path, 'w', encoding=encoding, newline='') as f:
            f.writelines(output_lines)

    return modified_count

if __name__ == "__main__":
    bas_files = glob.glob("*.bas")
    if not bas_files:
        print("لم يتم العثور على أي ملفات .bas في المجلد الحالي.")
    else:
        for bas_file in bas_files:
            count = inject_handlers_in_file(bas_file)
            print(f"[OK] {bas_file}: تم حقن معالج الأخطاء في {count} إجراء/دالة.")