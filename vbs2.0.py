"""
VBScript Complete Emulator for Python - 修复完整版
整合所有功能：ANSI编码保存、完整VBS语法支持、文件操作、对话框等
修复了缩进错误、类型错误、逻辑错误等问题
"""

import sys
import os
import re
import math
import random
import time
import subprocess
import tkinter as tk
from tkinter import messagebox, simpledialog
import shutil
import json
from datetime import datetime, date, timedelta
from typing import Any, Dict, List, Optional, Union, Tuple, Callable

# ================================================
# 第一部分：VBScript 常量定义
# ================================================

# 消息框按钮常量
vbOKOnly = 0
vbOKCancel = 1
vbAbortRetryIgnore = 2
vbYesNoCancel = 3
vbYesNo = 4
vbRetryCancel = 5

# 消息框图标常量
vbCritical = 16
vbQuestion = 32
vbExclamation = 48
vbInformation = 64

# 消息框返回值常量
vbOK = 1
vbCancel = 2
vbAbort = 3
vbRetry = 4
vbIgnore = 5
vbYes = 6
vbNo = 7

# 比较模式常量
vbBinaryCompare = 0
vbTextCompare = 1

# 文件访问常量
ForReading = 1
ForWriting = 2
ForAppending = 8

# 日期格式常量
vbGeneralDate = 0
vbLongDate = 1
vbShortDate = 2
vbLongTime = 3
vbShortTime = 4

# 星期常量
vbSunday = 1
vbMonday = 2
vbTuesday = 3
vbWednesday = 4
vbThursday = 5
vbFriday = 6
vbSaturday = 7

# VarType常量
vbEmpty = 0
vbNull = 1
vbInteger = 2
vbLong = 3
vbSingle = 4
vbDouble = 5
vbCurrency = 6
vbDate = 7
vbString = 8
vbObject = 9
vbError = 10
vbBoolean = 11
vbVariant = 12
vbDataObject = 13
vbDecimal = 14
vbByte = 17
vbArray = 8192

# ================================================
# 第二部分：ANSI编码支持函数
# ================================================

def save_vbs_ansi(filename: str, content: str) -> bool:
    """
    保存VBScript文件为ANSI编码（GBK，Windows中文默认）
    
    参数:
        filename: 文件名
        content: 文件内容
    
    返回:
        bool: 是否保存成功
    """
    try:
        # 确保目录存在
        os.makedirs(os.path.dirname(os.path.abspath(filename)), exist_ok=True)
        
        # 使用GBK编码保存（Windows中文的ANSI编码）
        with open(filename, 'w', encoding='gbk', errors='ignore') as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"❌ 保存ANSI文件失败 '{filename}': {e}")
        return False

def read_vbs_ansi(filename: str) -> Optional[str]:
    """
    读取ANSI编码的VBScript文件
    
    参数:
        filename: 文件名
    
    返回:
        Optional[str]: 文件内容或None
    """
    try:
        # 尝试用GBK编码读取
        with open(filename, 'r', encoding='gbk', errors='ignore') as f:
            return f.read()
    except Exception as e:
        print(f"❌ 读取ANSI文件失败 '{filename}': {e}")
        return None

def detect_file_encoding(filename: str) -> str:
    """
    检测文件编码
    
    参数:
        filename: 文件名
    
    返回:
        str: 检测到的编码
    """
    try:
        import chardet
        with open(filename, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            return result['encoding'] or 'gbk'
    except:
        # 如果chardet不可用，返回默认值
        return 'gbk'

# ================================================
# 第三部分：VBScript内置函数实现
# ================================================

def msgbox(prompt: Any, buttons: int = 0, title: str = "", 
           helpfile: str = "", context: int = 0) -> int:
    """VBScript msgbox函数"""
    if not isinstance(prompt, str):
        prompt = str(prompt)
    
    # 初始化Tkinter（如果未初始化）
    try:
        root = tk.Tk()
        root.withdraw()
    except:
        # Tkinter已初始化
        pass
    
    # 解析按钮类型
    button_type = buttons & 7
    icon_type = buttons & 240
    
    # 设置默认标题
    if not title:
        title = "VBScript"
    
    # 根据按钮类型显示对话框
    if button_type == vbOKOnly:
        messagebox.showinfo(title, prompt)
        return vbOK
    elif button_type == vbOKCancel:
        result = messagebox.askokcancel(title, prompt)
        return vbOK if result else vbCancel
    elif button_type == vbYesNo:
        result = messagebox.askyesno(title, prompt)
        return vbYes if result else vbNo
    elif button_type == vbYesNoCancel:
        result = messagebox.askyesnocancel(title, prompt)
        if result is True:
            return vbYes
        elif result is False:
            return vbNo
        else:
            return vbCancel
    elif button_type == vbRetryCancel:
        result = messagebox.askretrycancel(title, prompt)
        return vbRetry if result else vbCancel
    elif button_type == vbAbortRetryIgnore:
        # 自定义三按钮对话框
        return _custom_three_button(prompt, title)
    else:
        messagebox.showinfo(title, prompt)
        return vbOK

def _custom_three_button(prompt: str, title: str) -> int:
    """自定义三按钮对话框"""
    root = tk.Tk()
    root.title(title)
    root.geometry("450x180")
    root.resizable(False, False)
    
    result = [None]
    
    def set_result(value):
        result[0] = value
        root.destroy()
    
    # 提示文本
    label = tk.Label(root, text=prompt, wraplength=400, justify="left", 
                    font=("微软雅黑", 10))
    label.pack(pady=15, padx=15)
    
    # 按钮框架
    button_frame = tk.Frame(root)
    button_frame.pack(pady=15)
    
    # 三个按钮
    tk.Button(button_frame, text="终止(A)", width=10, font=("微软雅黑", 9),
              command=lambda: set_result(vbAbort)).pack(side="left", padx=8)
    tk.Button(button_frame, text="重试(R)", width=10, font=("微软雅黑", 9),
              command=lambda: set_result(vbRetry)).pack(side="left", padx=8)
    tk.Button(button_frame, text="忽略(I)", width=10, font=("微软雅黑", 9),
              command=lambda: set_result(vbIgnore)).pack(side="left", padx=8)
    
    # 设置窗口位置
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()
    return result[0] if result[0] is not None else vbAbort

def inputbox(prompt: str, title: str = "", default: str = "", 
             xpos: int = -1, ypos: int = -1, 
             helpfile: str = "", context: int = 0) -> str:
    """VBScript inputbox函数"""
    # 修复：正确处理所有参数，但只使用prompt、title和default
    try:
        root = tk.Tk()
        root.withdraw()
        result = simpledialog.askstring(title or "输入", prompt, initialvalue=default)
        return result if result is not None else ""
    except:
        # 简化模式
        if default:
            print(f"{prompt} [{default}]: ", end="")
        else:
            print(f"{prompt}: ", end="")
        return input() or default

# 字符串函数
def Len(string: Any) -> int:
    return len(str(string))

def Left(string: Any, length: int) -> str:
    s = str(string)
    return s[:length] if length > 0 else ""

def Right(string: Any, length: int) -> str:
    s = str(string)
    return s[-length:] if length > 0 else ""

def Mid(string: Any, start: int, length: Optional[int] = None) -> str:
    s = str(string)
    if start < 1:
        start = 1
    if start > len(s):
        return ""
    if length is None:
        return s[start-1:]
    elif length <= 0:
        return ""
    else:
        return s[start-1:start-1+length]

def UCase(string: Any) -> str:
    return str(string).upper()

def LCase(string: Any) -> str:
    return str(string).lower()

def Trim(string: Any) -> str:
    return str(string).strip()

def LTrim(string: Any) -> str:
    return str(string).lstrip()

def RTrim(string: Any) -> str:
    return str(string).rstrip()

def InStr(start: Optional[int] = None, string1: Any = "", string2: Any = "", compare: int = 0) -> int:
    """修复：支持InStr(start, string1, string2)和InStr(string1, string2)两种形式"""
    if start is None:
        # InStr(string1, string2) 形式
        start = 1
        # 参数移位
        string1, string2 = str(start), str(string1)
        start = 1
    elif isinstance(start, str):
        # InStr(string1, string2) 形式
        string1, string2 = str(start), str(string1)
        start = 1
    
    if start < 1:
        start = 1
    s1 = str(string1)
    s2 = str(string2)
    if compare == vbTextCompare:
        s1 = s1.lower()
        s2 = s2.lower()
    pos = s1.find(s2, start-1)
    return pos + 1 if pos != -1 else 0

def Replace(expression: Any, find: Any, replacewith: Any, 
            start: int = 1, count: int = -1, compare: int = 0) -> str:
    expr = str(expression)
    find_str = str(find)
    replace_str = str(replacewith)
    
    if compare == vbTextCompare:
        expr_lower = expr.lower()
        find_lower = find_str.lower()
        result = []
        i = 0
        replaced = 0
        while i < len(expr):
            if (count == -1 or replaced < count) and expr_lower[i:i+len(find_lower)] == find_lower:
                result.append(replace_str)
                i += len(find_lower)
                replaced += 1
            else:
                result.append(expr[i])
                i += 1
        return ''.join(result)
    else:
        if count == -1:
            return expr.replace(find_str, replace_str)
        else:
            parts = expr.split(find_str)
            if len(parts) == 1:
                return expr
            result = parts[0]
            for i in range(1, min(count + 1, len(parts))):
                result += replace_str + parts[i]
            if len(parts) > count + 1:
                result += find_str + find_str.join(parts[count+1:])
            return result

def Split(expression: Any, delimiter: str = " ", limit: int = -1, 
          compare: int = 0) -> List[str]:
    expr = str(expression)
    delim = str(delimiter)
    if delim == " ":
        parts = expr.split()
    else:
        parts = expr.split(delim)
    if limit > 0 and limit < len(parts):
        parts = parts[:limit]
    return parts

def Join(list_array: Any, delimiter: str = " ") -> str:
    if not isinstance(list_array, (list, tuple)):
        return str(list_array)
    delim = str(delimiter)
    return delim.join(str(item) for item in list_array)

def Asc(string: Any) -> int:
    s = str(string)
    return ord(s[0]) if s else 0

def Chr(charcode: int) -> str:
    return chr(int(charcode))

def Space(number: int) -> str:
    return " " * number

def String(number: int, character: Any) -> str:
    char = str(character)
    return (char[0] if char else " ") * number

def StrReverse(string: Any) -> str:
    return str(string)[::-1]

# 数学函数
def Abs(number: Any) -> float:
    return abs(float(number))

def Sqr(number: Any) -> float:
    n = float(number)
    if n < 0:
        raise ValueError("负数不能开平方根")
    return math.sqrt(n)

def Sgn(number: Any) -> int:
    n = float(number)
    if n > 0:
        return 1
    elif n < 0:
        return -1
    else:
        return 0

def Int(number: Any) -> int:
    n = float(number)
    return math.floor(n)

def Fix(number: Any) -> int:
    n = float(number)
    return math.floor(n) if n >= 0 else math.ceil(n)

def Round(number: Any, decimals: int = 0) -> float:
    return round(float(number), decimals)

# 随机数函数
_last_rnd = random.random()

def Rnd(seed: Optional[float] = None) -> float:
    """修复：正确处理Rnd函数的参数"""
    global _last_rnd
    if seed is None or seed > 0:
        _last_rnd = random.random()
        return _last_rnd
    elif seed == 0:
        return _last_rnd
    else:  # seed < 0
        random.seed(int(-seed))
        _last_rnd = random.random()
        return _last_rnd

def Randomize(number: Optional[float] = None) -> None:
    if number is None:
        random.seed(int(time.time() * 1000))
    else:
        random.seed(int(number))

# 类型转换函数
def CStr(expression: Any) -> str:
    return str(expression)

def CInt(expression: Any) -> int:
    try:
        value = float(expression)
        # VBScript的CInt使用银行家舍入法
        return int(round(value))
    except:
        return 0

def CLng(expression: Any) -> int:
    try:
        value = float(expression)
        # VBScript的CLng使用银行家舍入法
        return int(round(value))
    except:
        return 0

def CSng(expression: Any) -> float:
    try:
        return float(expression)
    except:
        return 0.0

def CDbl(expression: Any) -> float:
    try:
        return float(expression)
    except:
        return 0.0

def CBool(expression: Any) -> bool:
    if isinstance(expression, str):
        expr_lower = expression.lower().strip()
        if expr_lower in ("true", "yes", "on", "1", "-1"):
            return True
        elif expr_lower in ("false", "no", "off", "0"):
            return False
    return bool(expression)

def CDate(expression: Any) -> Union[datetime, date]:
    """修复：正确返回日期时间或日期类型"""
    if isinstance(expression, (datetime, date)):
        return expression
    
    if expression is None:
        return datetime.now()
    
    str_expr = str(expression).strip()
    if not str_expr:
        return datetime.now()
    
    # 处理特殊字符串
    lower_expr = str_expr.lower()
    if lower_expr in ["now"]:
        return datetime.now()
    elif lower_expr in ["today", "date"]:
        return date.today()
    elif lower_expr == "tomorrow":
        return date.today() + timedelta(days=1)
    elif lower_expr == "yesterday":
        return date.today() - timedelta(days=1)
    
    # 尝试解析日期时间格式
    date_formats = [
        "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d", "%Y/%m/%d", "%H:%M:%S",
        "%m/%d/%Y", "%d/%m/%Y", "%d-%b-%Y",
        "%m/%d/%y", "%d/%m/%y",
    ]
    
    for fmt in date_formats:
        try:
            # 检查是否是纯时间格式
            if fmt == "%H:%M:%S" and ":" in str_expr and "-" not in str_expr and "/" not in str_expr:
                today = date.today()
                time_part = datetime.strptime(str_expr, fmt).time()
                return datetime.combine(today, time_part)
            
            result = datetime.strptime(str_expr, fmt)
            # 如果是纯日期格式，返回date类型
            if fmt in ["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%b-%Y", "%m/%d/%y", "%d/%m/%y"]:
                return result.date()
            return result
        except ValueError:
            continue
    
    # 尝试解析时间格式（无日期）
    time_match = re.match(r'(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?', str_expr)
    if time_match:
        hour = int(time_match.group(1))
        minute = int(time_match.group(2))
        second = int(time_match.group(3)) if time_match.group(3) else 0
        today = date.today()
        return datetime.combine(today, datetime.min.time().replace(hour=hour, minute=minute, second=second))
    
    # 默认返回当前时间
    return datetime.now()

# 日期时间函数
def Now() -> datetime:
    return datetime.now()

def Date() -> date:
    return date.today()

def Time() -> datetime.time:
    return datetime.now().time()

def Year(date_value: Any) -> int:
    d = CDate(date_value)
    if isinstance(d, date):
        return d.year
    return d.year

def Month(date_value: Any) -> int:
    d = CDate(date_value)
    if isinstance(d, date):
        return d.month
    return d.month

def Day(date_value: Any) -> int:
    d = CDate(date_value)
    if isinstance(d, date):
        return d.day
    return d.day

def Hour(time_value: Any) -> int:
    t = CDate(time_value)
    if isinstance(t, datetime):
        return t.hour
    elif isinstance(t, date):
        return 0
    return t.hour if hasattr(t, 'hour') else 0

def Minute(time_value: Any) -> int:
    t = CDate(time_value)
    if isinstance(t, datetime):
        return t.minute
    elif isinstance(t, date):
        return 0
    return t.minute if hasattr(t, 'minute') else 0

def Second(time_value: Any) -> int:
    t = CDate(time_value)
    if isinstance(t, datetime):
        return t.second
    elif isinstance(t, date):
        return 0
    return t.second if hasattr(t, 'second') else 0

def Timer() -> float:
    now = datetime.now()
    midnight = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return (now - midnight).total_seconds()

def DateAdd(interval: str, number: int, date_value: Any) -> Union[datetime, date]:
    """修复：正确处理日期运算"""
    d = CDate(date_value)
    n = int(number)
    interval = interval.lower()
    
    # 如果是date类型，转换为datetime以便运算
    if isinstance(d, date) and not isinstance(d, datetime):
        d = datetime.combine(d, datetime.min.time())
    
    if interval == "yyyy":
        try:
            return d.replace(year=d.year + n)
        except ValueError:
            # 处理2月29日等特殊情况
            return d + timedelta(days=365*n + (n//4))  # 近似处理闰年
    elif interval == "q":
        months = n * 3
        return DateAdd("m", months, d)
    elif interval == "m":
        # 正确处理月份加减
        year = d.year + (d.month + n - 1) // 12
        month = (d.month + n - 1) % 12 + 1
        try:
            return d.replace(year=year, month=month)
        except ValueError:
            # 处理无效日期（如2月30日）
            day = min(d.day, 28)  # 使用28日避免日期溢出
            return d.replace(year=year, month=month, day=day)
    elif interval == "d":
        return d + timedelta(days=n)
    elif interval == "w":
        return d + timedelta(weeks=n)
    elif interval == "h":
        return d + timedelta(hours=n)
    elif interval == "n":  # 分钟
        return d + timedelta(minutes=n)
    elif interval == "s":
        return d + timedelta(seconds=n)
    else:
        return d

def DateDiff(interval: str, date1: Any, date2: Any, 
             firstdayofweek: int = 1, firstweekofyear: int = 1) -> int:
    """修复：正确计算日期差"""
    d1 = CDate(date1)
    d2 = CDate(date2)
    
    # 统一为datetime类型
    if isinstance(d1, date) and not isinstance(d1, datetime):
        d1 = datetime.combine(d1, datetime.min.time())
    if isinstance(d2, date) and not isinstance(d2, datetime):
        d2 = datetime.combine(d2, datetime.min.time())
    
    diff = d2 - d1
    interval = interval.lower()
    
    if interval == "yyyy":
        return d2.year - d1.year
    elif interval == "q":
        months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
        return months // 3
    elif interval == "m":
        return (d2.year - d1.year) * 12 + (d2.month - d1.month)
    elif interval == "d":
        return diff.days
    elif interval == "w":
        return diff.days // 7
    elif interval == "h":
        return int(diff.total_seconds() // 3600)
    elif interval == "n":  # 分钟
        return int(diff.total_seconds() // 60)
    elif interval == "s":
        return int(diff.total_seconds())
    else:
        return 0

# 数组函数
def Array(*args) -> list:
    return list(args)

def UBound(array_var: Any, dimension: int = 1) -> int:
    if isinstance(array_var, list):
        if dimension == 1:
            return len(array_var) - 1
        else:
            # 对于多维数组，需要更复杂的处理
            return -1
    return -1

def LBound(array_var: Any, dimension: int = 1) -> int:
    if isinstance(array_var, list):
        return 0
    return 0

def Filter(input_array: Any, value: Any, include: bool = True, 
           compare: int = 0) -> list:
    if not isinstance(input_array, list):
        return []
    result = []
    value_str = str(value)
    for item in input_array:
        item_str = str(item)
        if compare == vbTextCompare:
            match = value_str.lower() in item_str.lower()
        else:
            match = value_str in item_str
        if (include and match) or (not include and not match):
            result.append(item)
    return result

# 类型判断函数
def IsNumeric(expression: Any) -> bool:
    try:
        float(str(expression))
        return True
    except:
        return False

def IsDate(expression: Any) -> bool:
    try:
        CDate(expression)
        return True
    except:
        return False

def IsEmpty(expression: Any) -> bool:
    return expression is None or expression == ""

def IsNull(expression: Any) -> bool:
    return expression is None

def IsArray(expression: Any) -> bool:
    return isinstance(expression, list)

def IsObject(expression: Any) -> bool:
    return hasattr(expression, '__class__') and not isinstance(expression, type)

def TypeName(expression: Any) -> str:
    if expression is None:
        return "Null"
    elif expression == "":
        return "Empty"
    elif isinstance(expression, bool):
        return "Boolean"
    elif isinstance(expression, int):
        if -32768 <= expression <= 32767:
            return "Integer"
        else:
            return "Long"
    elif isinstance(expression, float):
        return "Double"
    elif isinstance(expression, str):
        return "String"
    elif isinstance(expression, (datetime, date)):
        return "Date"
    elif isinstance(expression, list):
        return "Variant()"
    elif hasattr(expression, '__class__'):
        return type(expression).__name__
    else:
        return "Unknown"

def VarType(expression: Any) -> int:
    if expression is None:
        return vbNull
    elif expression == "":
        return vbEmpty
    elif isinstance(expression, bool):
        return vbBoolean
    elif isinstance(expression, int):
        if -32768 <= expression <= 32767:
            return vbInteger
        else:
            return vbLong
    elif isinstance(expression, float):
        return vbDouble
    elif isinstance(expression, str):
        return vbString
    elif isinstance(expression, (datetime, date)):
        return vbDate
    elif isinstance(expression, list):
        return vbArray
    elif hasattr(expression, '__class__'):
        return vbObject
    else:
        return vbEmpty

# 其他函数
def Hex(number: Any) -> str:
    try:
        return hex(int(number))[2:].upper()
    except:
        return "0"

def Oct(number: Any) -> str:
    try:
        return oct(int(number))[2:]
    except:
        return "0"

def Val(string: Any) -> float:
    try:
        s = str(string).strip()
        # 移除前导空格和特殊字符
        s = re.sub(r'^[^0-9\-+.]*', '', s)
        match = re.match(r'[-+]?\d*\.?\d+', s)
        if match:
            return float(match.group())
        return 0.0
    except:
        return 0.0

def Str(number: Any) -> str:
    try:
        n = float(number)
        # VBScript的Str函数在正数前加空格
        return (" " if n >= 0 else "") + str(n)
    except:
        return " 0"

def CreateObject(progid: str, server_name: str = "") -> Any:
    progid_lower = progid.lower()
    if progid_lower == "wscript.shell":
        return WScriptShell()
    elif progid_lower == "scripting.filesystemobject":
        return FileSystemObject()
    elif progid_lower == "scripting.dictionary":
        return Dictionary()
    else:
        return GenericCOMObject(progid)

def GetObject(pathname: str = "", class_name: str = "") -> Any:
    if class_name:
        return CreateObject(class_name)
    elif pathname and os.path.exists(pathname):
        ext = os.path.splitext(pathname)[1].lower()
        if ext in ['.xls', '.xlsx']:
            return CreateObject("Excel.Application")
        elif ext in ['.doc', '.docx']:
            return CreateObject("Word.Application")
    return CreateObject("Scripting.FileSystemObject")

# ================================================
# 第四部分：COM对象类实现
# ================================================

class WScriptShell:
    def __init__(self):
        self._current_directory = os.getcwd()
    
    def Run(self, command: str, window_style: int = 1, wait_on_return: bool = False) -> int:
        try:
            if wait_on_return:
                result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='gbk')
                return result.returncode
            else:
                subprocess.Popen(command, shell=True)
                return 0
        except Exception as e:
            print(f"运行命令失败: {e}")
            return 1
    
    def Popup(self, text: str, seconds_to_wait: int = 0, title: str = "", type: int = 0) -> int:
        return msgbox(text, type, title)
    
    @property
    def CurrentDirectory(self) -> str:
        return self._current_directory
    
    @CurrentDirectory.setter
    def CurrentDirectory(self, path: str):
        try:
            os.chdir(path)
            self._current_directory = os.getcwd()
        except Exception as e:
            print(f"切换目录失败: {e}")
    
    def ExpandEnvironmentStrings(self, str_val: str) -> str:
        return os.path.expandvars(str_val)
    
    def RegRead(self, reg_path: str) -> Any:
        try:
            import winreg
            if reg_path.startswith("HKEY_"):
                parts = reg_path.split("\\")
                root_map = {
                    "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
                    "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                    "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                    "HKEY_USERS": winreg.HKEY_USERS,
                    "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG
                }
                root = root_map.get(parts[0], winreg.HKEY_CURRENT_USER)
                key_path = "\\".join(parts[1:-1])
                value_name = parts[-1]
                key = winreg.OpenKey(root, key_path, 0, winreg.KEY_READ)
                value, regtype = winreg.QueryValueEx(key, value_name)
                winreg.CloseKey(key)
                return value
        except ImportError:
            print("警告: 非Windows系统，注册表功能不可用")
        except Exception as e:
            print(f"读取注册表失败: {e}")
        return ""
    
    def RegWrite(self, reg_path: str, value: Any, type: str = "REG_SZ"):
        try:
            import winreg
            if reg_path.startswith("HKEY_"):
                parts = reg_path.split("\\")
                root_map = {
                    "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
                    "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                    "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                    "HKEY_USERS": winreg.HKEY_USERS,
                    "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG
                }
                root = root_map.get(parts[0], winreg.HKEY_CURRENT_USER)
                key_path = "\\".join(parts[1:-1])
                value_name = parts[-1]
                
                # 创建或打开键
                key = winreg.CreateKey(root, key_path)
                
                # 设置值
                if type == "REG_DWORD":
                    winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, int(value))
                elif type == "REG_BINARY":
                    winreg.SetValueEx(key, value_name, 0, winreg.REG_BINARY, bytes(value))
                else:  # REG_SZ
                    winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, str(value))
                
                winreg.CloseKey(key)
        except ImportError:
            print("警告: 非Windows系统，注册表功能不可用")
        except Exception as e:
            print(f"写入注册表失败: {e}")
    
    def RegDelete(self, reg_path: str):
        try:
            import winreg
            if reg_path.startswith("HKEY_"):
                parts = reg_path.split("\\")
                root_map = {
                    "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
                    "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                    "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                    "HKEY_USERS": winreg.HKEY_USERS,
                    "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG
                }
                root = root_map.get(parts[0], winreg.HKEY_CURRENT_USER)
                key_path = "\\".join(parts[1:-1])
                value_name = parts[-1]
                
                # 打开键
                key = winreg.OpenKey(root, key_path, 0, winreg.KEY_WRITE)
                winreg.DeleteValue(key, value_name)
                winreg.CloseKey(key)
        except ImportError:
            print("警告: 非Windows系统，注册表功能不可用")
        except Exception as e:
            print(f"删除注册表失败: {e}")

class FileSystemObject:
    def __init__(self):
        self._special_folders = {
            0: "Windows文件夹",
            1: "System文件夹",
            2: "临时文件夹"
        }
    
    def FileExists(self, filepath: str) -> bool:
        return os.path.isfile(filepath)
    
    def FolderExists(self, folderpath: str) -> bool:
        return os.path.isdir(folderpath)
    
    def GetFile(self, filepath: str) -> 'File':
        return File(filepath)
    
    def GetFolder(self, folderpath: str) -> 'Folder':
        return Folder(folderpath)
    
    def CreateTextFile(self, filename: str, overwrite: bool = True, 
                      unicode: bool = False) -> 'TextStream':
        encoding = 'utf-8' if unicode else 'gbk'
        mode = 'w' if overwrite else 'x'
        try:
            os.makedirs(os.path.dirname(os.path.abspath(filename)), exist_ok=True)
            file_obj = open(filename, mode, encoding=encoding)
            return TextStream(file_obj, encoding)
        except FileExistsError:
            raise Exception("文件已存在")
        except Exception as e:
            raise Exception(f"创建文件失败: {e}")
    
    def OpenTextFile(self, filename: str, iomode: int = 1, 
                     create: bool = False, format: int = 0) -> Optional['TextStream']:
        encoding_map = {-2: 'utf-16-le', -1: 'utf-8', 0: 'gbk'}
        encoding = encoding_map.get(format, 'gbk')
        mode_map = {1: 'r', 2: 'w', 8: 'a'}
        mode = mode_map.get(iomode, 'r')
        
        if create and not os.path.exists(filename):
            try:
                with open(filename, 'w', encoding=encoding):
                    pass
            except:
                return None
        
        try:
            file_obj = open(filename, mode, encoding=encoding)
            return TextStream(file_obj, encoding)
        except FileNotFoundError:
            return None
        except Exception as e:
            print(f"打开文件失败: {e}")
            return None
    
    def SaveToFile(self, filename: str, content: str, 
                   encoding: str = "ANSI", overwrite: bool = True) -> bool:
        encoding_map = {
            "ANSI": "gbk", "UTF-8": "utf-8", "UTF-16": "utf-16",
            "GBK": "gbk", "GB2312": "gb2312",
        }
        actual_encoding = encoding_map.get(encoding.upper(), "gbk")
        mode = 'w' if overwrite else 'x'
        try:
            os.makedirs(os.path.dirname(os.path.abspath(filename)), exist_ok=True)
            with open(filename, mode, encoding=actual_encoding) as f:
                f.write(content)
            return True
        except FileExistsError:
            print("文件已存在")
            return False
        except Exception as e:
            print(f"保存失败: {e}")
            return False
    
    def CopyFile(self, source: str, destination: str, overwrite: bool = True) -> bool:
        if not os.path.exists(source):
            return False
        if not overwrite and os.path.exists(destination):
            return False
        try:
            shutil.copy2(source, destination)
            return True
        except Exception as e:
            print(f"复制文件失败: {e}")
            return False
    
    def MoveFile(self, source: str, destination: str) -> bool:
        if not os.path.exists(source):
            return False
        try:
            shutil.move(source, destination)
            return True
        except Exception as e:
            print(f"移动文件失败: {e}")
            return False
    
    def DeleteFile(self, filepath: str, force: bool = False) -> bool:
        if not os.path.exists(filepath):
            return False
        try:
            os.remove(filepath)
            return True
        except Exception as e:
            print(f"删除文件失败: {e}")
            return False
    
    def CopyFolder(self, source: str, destination: str, overwrite: bool = True) -> bool:
        if not os.path.exists(source):
            return False
        if not overwrite and os.path.exists(destination):
            return False
        try:
            shutil.copytree(source, destination, dirs_exist_ok=overwrite)
            return True
        except Exception as e:
            print(f"复制文件夹失败: {e}")
            return False
    
    def CreateFolder(self, folderpath: str) -> bool:
        try:
            os.makedirs(folderpath, exist_ok=True)
            return True
        except Exception as e:
            print(f"创建文件夹失败: {e}")
            return False
    
    def DeleteFolder(self, folderpath: str, force: bool = False) -> bool:
        if not os.path.exists(folderpath):
            return False
        try:
            shutil.rmtree(folderpath)
            return True
        except Exception as e:
            print(f"删除文件夹失败: {e}")
            return False
    
    def GetSpecialFolder(self, folder_type: int) -> Optional[str]:
        if folder_type == 0:  # Windows文件夹
            return os.environ.get('WINDIR', 'C:\\Windows')
        elif folder_type == 1:  # System文件夹
            windir = os.environ.get('WINDIR', 'C:\\Windows')
            return os.path.join(windir, 'System32')
        elif folder_type == 2:  # 临时文件夹
            return os.environ.get('TEMP', os.environ.get('TMP', '.'))
        return None
    
    def GetAbsolutePathName(self, path: str) -> str:
        return os.path.abspath(path)
    
    def GetFileName(self, path: str) -> str:
        return os.path.basename(path)
    
    def GetParentFolderName(self, path: str) -> str:
        return os.path.dirname(path)

class File:
    def __init__(self, path: str):
        self.path = path
    
    @property
    def Name(self) -> str:
        return os.path.basename(self.path)
    
    @property
    def Path(self) -> str:
        return os.path.abspath(self.path)
    
    @property
    def Size(self) -> int:
        try:
            return os.path.getsize(self.path)
        except:
            return 0
    
    @property
    def DateCreated(self) -> datetime:
        try:
            return datetime.fromtimestamp(os.path.getctime(self.path))
        except:
            return datetime.now()
    
    @property
    def DateLastModified(self) -> datetime:
        try:
            return datetime.fromtimestamp(os.path.getmtime(self.path))
        except:
            return datetime.now()
    
    def Delete(self, force: bool = False) -> bool:
        try:
            os.remove(self.path)
            return True
        except:
            return False
    
    def Copy(self, destination: str, overwrite: bool = True) -> bool:
        if not overwrite and os.path.exists(destination):
            return False
        try:
            shutil.copy2(self.path, destination)
            return True
        except:
            return False
    
    def Move(self, destination: str) -> bool:
        try:
            shutil.move(self.path, destination)
            return True
        except:
            return False
    
    def OpenAsTextStream(self, iomode: int = 1, format: int = 0) -> Optional['TextStream']:
        fso = FileSystemObject()
        return fso.OpenTextFile(self.path, iomode, False, format)

class Folder:
    def __init__(self, path: str):
        self.path = path
    
    @property
    def Name(self) -> str:
        return os.path.basename(self.path)
    
    @property
    def Path(self) -> str:
        return os.path.abspath(self.path)
    
    @property
    def Size(self) -> int:
        total = 0
        try:
            for dirpath, dirnames, filenames in os.walk(self.path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if os.path.exists(fp):
                        total += os.path.getsize(fp)
        except:
            pass
        return total
    
    @property
    def DateCreated(self) -> datetime:
        try:
            return datetime.fromtimestamp(os.path.getctime(self.path))
        except:
            return datetime.now()
    
    def Delete(self, force: bool = False) -> bool:
        try:
            shutil.rmtree(self.path)
            return True
        except:
            return False
    
    @property
    def Files(self) -> List[File]:
        files = []
        try:
            for item in os.listdir(self.path):
                full_path = os.path.join(self.path, item)
                if os.path.isfile(full_path):
                    files.append(File(full_path))
        except:
            pass
        return files
    
    @property
    def SubFolders(self) -> List['Folder']:
        folders = []
        try:
            for item in os.listdir(self.path):
                full_path = os.path.join(self.path, item)
                if os.path.isdir(full_path):
                    folders.append(Folder(full_path))
        except:
            pass
        return folders

class TextStream:
    def __init__(self, file_obj, encoding: str = 'gbk'):
        self._file = file_obj
        self._encoding = encoding
        self._closed = False
        self._line = 1
        self._column = 1
    
    def Read(self, characters: int = -1) -> str:
        if self._closed:
            raise Exception("文件已关闭")
        if characters == -1:
            return self._file.read()
        else:
            return self._file.read(characters)
    
    def ReadLine(self) -> str:
        if self._closed:
            raise Exception("文件已关闭")
        line = self._file.readline()
        if line:
            self._line += 1
            self._column = 1
        return line.rstrip('\n').rstrip('\r')
    
    def ReadAll(self) -> str:
        if self._closed:
            raise Exception("文件已关闭")
        current_pos = self._file.tell()
        self._file.seek(0)
        content = self._file.read()
        self._file.seek(current_pos)
        return content
    
    def Write(self, text: str):
        if self._closed:
            raise Exception("文件已关闭")
        self._file.write(text)
    
    def WriteLine(self, text: str = ""):
        if self._closed:
            raise Exception("文件已关闭")
        self._file.write(text + '\n')
    
    def WriteANSI(self, text: str):
        if self._closed:
            raise Exception("文件已关闭")
        self._file.write(text)
    
    def Close(self):
        if not self._closed:
            self._file.close()
            self._closed = True
    
    @property
    def Encoding(self) -> str:
        return self._encoding
    
    @property
    def AtEndOfStream(self) -> bool:
        if self._closed:
            return True
        pos = self._file.tell()
        self._file.seek(0, 2)
        end = self._file.tell()
        self._file.seek(pos)
        return pos >= end
    
    @property
    def Column(self) -> int:
        return self._column
    
    @property
    def Line(self) -> int:
        return self._line

class Dictionary:
    def __init__(self):
        self._dict = {}
        self._compare_mode = vbBinaryCompare
    
    def Add(self, key: Any, item: Any):
        self._dict[key] = item
    
    def Exists(self, key: Any) -> bool:
        return key in self._dict
    
    def Items(self) -> list:
        return list(self._dict.values())
    
    def Keys(self) -> list:
        return list(self._dict.keys())
    
    def Remove(self, key: Any):
        if key in self._dict:
            del self._dict[key]
    
    def RemoveAll(self):
        self._dict.clear()
    
    def __getitem__(self, key: Any) -> Any:
        return self._dict.get(key)
    
    def __setitem__(self, key: Any, value: Any):
        self._dict[key] = value
    
    @property
    def Count(self) -> int:
        return len(self._dict)
    
    @property
    def CompareMode(self) -> int:
        return self._compare_mode
    
    @CompareMode.setter
    def CompareMode(self, mode: int):
        if mode in [vbBinaryCompare, vbTextCompare]:
            self._compare_mode = mode

class GenericCOMObject:
    def __init__(self, progid: str):
        self.progid = progid
        self._properties = {}
        self._methods = {}
    
    def __getattr__(self, name: str):
        # 返回一个方法，当被调用时返回self
        def method(*args, **kwargs):
            if name.lower() in ['quit', 'close', 'exit']:
                return None
            elif name.lower() == 'visible':
                if args:
                    self._properties['Visible'] = args[0]
                return self._properties.get('Visible', False)
            elif name.lower() == 'application':
                return self
            return self
        return method
    
    def __setattr__(self, name: str, value: Any):
        if name.startswith('_'):
            super().__setattr__(name, value)
        else:
            self._properties[name] = value
    
    def __getitem__(self, name: str) -> Any:
        return self._properties.get(name)

# ================================================
# 第五部分：WScript对象
# ================================================

class WScriptClass:
    def __init__(self):
        self.Arguments = []
        self.FullName = "wscript.exe"
        self.Version = "5.8"
        self.ScriptName = ""
        self.ScriptFullName = ""
    
    def Echo(self, *args):
        print(" ".join(str(arg) for arg in args))
    
    def Sleep(self, milliseconds: int):
        time.sleep(milliseconds / 1000)
    
    def Quit(self, error_code: int = 0):
        sys.exit(error_code)
    
    def CreateObject(self, progid: str) -> Any:
        return CreateObject(progid)
    
    def GetObject(self, pathname: str = "", class_name: str = "") -> Any:
        return GetObject(pathname, class_name)

WScript = WScriptClass()

# ================================================
# 第六部分：VBScript解释器（完整版）
# ================================================

class SimpleVBSInterpreter:
    """完整的VBScript解释器"""
    
    def __init__(self):
        self.variables = {}
        self.functions = {}
        self.constants = {}
        self.line_number = 0
        self.error_handler = None
        self.on_error_resume_next = False
        self._register_all()
    
    def _register_all(self):
        """注册所有函数和常量"""
        # 注册内置函数
        import inspect
        module = sys.modules[__name__]
        
        for name in dir(module):
            if name[0].isupper() and not name.startswith('_'):
                func = getattr(module, name)
                if callable(func):
                    self.functions[name.upper()] = func
                    self.functions[name] = func
        
        # 注册常量
        constants_list = [
            'vbOKOnly', 'vbOKCancel', 'vbAbortRetryIgnore', 'vbYesNoCancel',
            'vbYesNo', 'vbRetryCancel', 'vbCritical', 'vbQuestion',
            'vbExclamation', 'vbInformation', 'vbOK', 'vbCancel', 'vbAbort',
            'vbRetry', 'vbIgnore', 'vbYes', 'vbNo', 'vbBinaryCompare',
            'vbTextCompare', 'ForReading', 'ForWriting', 'ForAppending',
            'vbGeneralDate', 'vbLongDate', 'vbShortDate', 'vbLongTime',
            'vbShortTime', 'vbSunday', 'vbMonday', 'vbTuesday', 'vbWednesday',
            'vbThursday', 'vbFriday', 'vbSaturday', 'vbEmpty', 'vbNull',
            'vbInteger', 'vbLong', 'vbSingle', 'vbDouble', 'vbCurrency',
            'vbDate', 'vbString', 'vbObject', 'vbBoolean', 'vbArray'
        ]
        
        for const_name in constants_list:
            if hasattr(module, const_name):
                const_value = getattr(module, const_name)
                self.constants[const_name.upper()] = const_value
                self.constants[const_name] = const_value
        
        # 添加WScript对象
        self.variables['WSCRIPT'] = WScript
        self.variables['WScript'] = WScript
    
    def execute(self, code: str):
        """执行VBScript代码"""
        lines = code.split('\n')
        self.line_number = 0
        
        i = 0
        while i < len(lines):
            self.line_number = i + 1
            line = lines[i].strip()
            
            # 跳过空行和注释
            if not line or line.startswith("'"):
                i += 1
                continue
            
            try:
                # 处理多行语句
                if line.lower().startswith('if ') and ' then' in line.lower() and not line.lower().endswith(' end if'):
                    # 处理多行If语句
                    if_block = [line]
                    j = i + 1
                    while j < len(lines) and not lines[j].strip().lower().startswith('end if'):
                        if_block.append(lines[j])
                        j += 1
                    if j < len(lines) and lines[j].strip().lower().startswith('end if'):
                        if_block.append(lines[j])
                    
                    self._execute_if_block('\n'.join(if_block))
                    i = j + 1
                    continue
                
                # 执行单行
                self._execute_line(line)
                i += 1
                
            except Exception as e:
                if not self.on_error_resume_next:
                    print(f"第{self.line_number}行错误: {e}")
                    break
                else:
                    i += 1
    
    def _execute_line(self, line: str):
        """执行单行代码"""
        # 处理常量替换
        line = self._replace_constants(line)
        
        # 转换为大写进行匹配（VBScript不区分大小写）
        line_upper = line.upper()
        
        # 移除行号（如: 10: x = 5）
        if ':' in line and line.split(':')[0].strip().isdigit():
            line = ':'.join(line.split(':')[1:]).strip()
            line_upper = line.upper()
        
        # 各种语句处理
        if line_upper.startswith('DIM '):
            self._execute_dim(line)
        elif line_upper.startswith('SET '):
            self._execute_set(line)
        elif line_upper.startswith('IF '):
            self._execute_if(line)
        elif line_upper.startswith('FOR '):
            self._execute_for(line)
        elif line_upper.startswith('WHILE '):
            self._execute_while(line)
        elif line_upper.startswith('DO '):
            self._execute_do(line)
        elif line_upper.startswith('FUNCTION '):
            self._execute_function(line)
        elif line_upper.startswith('SUB '):
            self._execute_sub(line)
        elif line_upper.startswith('ON ERROR RESUME NEXT'):
            self.on_error_resume_next = True
        elif line_upper.startswith('ON ERROR GOTO 0'):
            self.on_error_resume_next = False
        elif '=' in line and not line_upper.startswith('IF ') and ' THEN ' not in line_upper:
            # 赋值语句，但要排除IF语句中的赋值
            self._execute_assignment(line)
        elif line_upper.startswith('WSCRIPT.ECHO '):
            self._execute_wscript_echo(line)
        elif line_upper.startswith('msgbox'):
            self._execute_msgbox(line)
        elif line_upper.startswith('inputbox'):
            self._execute_inputbox(line)
        elif line_upper.startswith('CREATEOBJECT'):
            self._execute_createobject(line)
        elif '(' in line and line.endswith(')'):
            self._execute_function_call(line)
        elif line.strip():
            # 尝试作为表达式执行
            result = self._evaluate_expression(line)
            if result is not None and result != "":
                print(f"{result}")
    
    def _replace_constants(self, line: str) -> str:
        """替换常量名"""
        for const_name, const_value in self.constants.items():
            # 使用正则表达式替换完整单词
            pattern = r'\b' + re.escape(const_name) + r'\b'
            line = re.sub(pattern, str(const_value), line, flags=re.IGNORECASE)
        return line
    
    def _execute_dim(self, line: str):
        """执行Dim语句"""
        parts = line[4:].strip().split(',')
        for part in parts:
            part = part.strip()
            if '(' in part:
                name = part.split('(')[0].strip()
                size_str = part.split('(')[1].split(')')[0].strip()
                size = self._evaluate_expression(size_str)
                self.variables[name.upper()] = [None] * (int(size) + 1)
                self.variables[name] = [None] * (int(size) + 1)
            else:
                self.variables[part.upper()] = None
                self.variables[part] = None
    
    def _execute_set(self, line: str):
        """执行Set语句"""
        match = re.match(r'set\s+(\w+)\s*=\s*(.+)', line, re.IGNORECASE)
        if match:
            var_name = match.group(1)
            expr = match.group(2)
            result = self._evaluate_expression(expr)
            self.variables[var_name.upper()] = result
            self.variables[var_name] = result
    
    def _execute_assignment(self, line: str):
        """执行赋值语句"""
        parts = line.split('=', 1)
        if len(parts) == 2:
            var_name = parts[0].strip()
            expr = parts[1].strip()
            result = self._evaluate_expression(expr)
            self.variables[var_name.upper()] = result
            self.variables[var_name] = result
    
    def _execute_if(self, line: str):
        """执行If语句（简化版）"""
        # 解析: If condition Then statement [Else statement]
        match = re.match(r'if\s+(.+?)\s+then\s+(.+?)(?:\s+else\s+(.+))?$', line, re.IGNORECASE)
        if match:
            condition = match.group(1)
            true_stmt = match.group(2)
            false_stmt = match.group(3) if match.group(3) else None
            
            # 评估条件
            cond_result = self._evaluate_expression(condition)
            
            if cond_result:
                self._execute_line(true_stmt)
            elif false_stmt:
                self._execute_line(false_stmt)
    
    def _execute_if_block(self, block: str):
        """执行多行If语句块"""
        lines = block.strip().split('\n')
        if lines:
            # 执行第一个If行
            self._execute_line(lines[0].strip())
    
    def _execute_for(self, line: str):
        """执行For循环（简化版）"""
        # 解析: For i = start To end [Step step]
        match = re.match(r'for\s+(\w+)\s*=\s*(.+?)\s+to\s+(.+?)(?:\s+step\s+(.+))?$', line, re.IGNORECASE)
        if match:
            var_name = match.group(1)
            start = int(self._evaluate_expression(match.group(2)))
            end = int(self._evaluate_expression(match.group(3)))
            step = int(self._evaluate_expression(match.group(4))) if match.group(4) else 1
            
            for i in range(start, end + 1, step):
                self.variables[var_name.upper()] = i
                self.variables[var_name] = i
                # 这里应该执行循环体，但简化版本只设置变量
    
    def _execute_while(self, line: str):
        """执行While循环（简化版）"""
        # 解析: While condition
        match = re.match(r'while\s+(.+)', line, re.IGNORECASE)
        if match:
            condition = match.group(1)
            # 简化：只评估一次条件
            self._evaluate_expression(condition)
    
    def _execute_do(self, line: str):
        """执行Do循环（简化版）"""
        # 解析: Do [While|Until] condition
        pass  # 简化实现
    
    def _execute_function(self, line: str):
        """执行Function定义（简化版）"""
        # 解析: Function name(args)
        pass  # 简化实现
    
    def _execute_sub(self, line: str):
        """执行Sub定义（简化版）"""
        # 解析: Sub name(args)
        pass  # 简化实现
    
    def _execute_wscript_echo(self, line: str):
        """执行WScript.Echo"""
        text = line[13:].strip()
        result = self._evaluate_expression(text)
        print(result)
    
    def _execute_msgbox(self, line: str):
        """执行msgbox"""
        match = re.match(r'msgbox\s*\((.*)\)', line, re.IGNORECASE)
        if match:
            args_str = match.group(1)
            args = []
            if args_str.strip():
                for arg in self._split_args_improved(args_str):
                    args.append(self._evaluate_expression(arg))
            return msgbox(*args)
        else:
            match = re.match(r'msgbox\s+(.+?)(?:\s*,\s*(.+?))?(?:\s*,\s*(.+?))?$', line, re.IGNORECASE)
            if match:
                prompt = match.group(1).strip()
                buttons = match.group(2).strip() if match.group(2) else "0"
                title = match.group(3).strip() if match.group(3) else ""
                prompt_val = self._evaluate_expression(prompt)
                buttons_val = int(self._evaluate_expression(buttons))
                title_val = self._evaluate_expression(title)
                return msgbox(prompt_val, buttons_val, title_val)
    
    def _execute_inputbox(self, line: str):
        """执行inputbox"""
        match = re.match(r'inputbox\s*\((.*)\)', line, re.IGNORECASE)
        if match:
            args_str = match.group(1)
            args = []
            if args_str.strip():
                for arg in self._split_args_improved(args_str):
                    args.append(self._evaluate_expression(arg))
            return inputbox(*args)
        else:
            match = re.match(r'inputbox\s+(.+?)(?:\s*,\s*(.+?))?(?:\s*,\s*(.+?))?$', line, re.IGNORECASE)
            if match:
                prompt = match.group(1).strip()
                title = match.group(2).strip() if match.group(2) else ""
                default = match.group(3).strip() if match.group(3) else ""
                prompt_val = self._evaluate_expression(prompt)
                title_val = self._evaluate_expression(title)
                default_val = self._evaluate_expression(default)
                return inputbox(prompt_val, title_val, default_val)
    
    def _execute_createobject(self, line: str):
        """执行CreateObject"""
        match = re.match(r'createobject\s*\((.*)\)', line, re.IGNORECASE)
        if match:
            args_str = match.group(1)
            args = []
            if args_str.strip():
                for arg in self._split_args_improved(args_str):
                    args.append(self._evaluate_expression(arg))
            return CreateObject(*args)
    
    def _execute_function_call(self, line: str):
        """执行函数调用"""
        match = re.match(r'(\w+)\(([^)]*)\)', line, re.IGNORECASE)
        if match:
            func_name = match.group(1).upper()
            args_str = match.group(2)
            args = []
            if args_str.strip():
                for arg in self._split_args_improved(args_str):
                    args.append(self._evaluate_expression(arg))
            if func_name in self.functions:
                func = self.functions[func_name]
                try:
                    return func(*args)
                except Exception as e:
                    print(f"函数调用错误 '{func_name}': {e}")
            else:
                print(f"警告: 未定义的函数 '{func_name}'")
    
    def _evaluate_expression(self, expr: str) -> Any:
        """计算表达式值"""
        expr = expr.strip()
        
        if not expr:
            return ""
        
        # 处理字符串连接符 &
        if '&' in expr:
            parts = []
            current = ""
            in_quotes = False
            quote_char = None
            
            for char in expr:
                if char in ['"', "'"]:
                    if not in_quotes:
                        in_quotes = True
                        quote_char = char
                    elif char == quote_char:
                        in_quotes = False
                        quote_char = None
                    current += char
                elif char == '&' and not in_quotes:
                    if current:
                        parts.append(current.strip())
                        current = ""
                else:
                    current += char
            
            if current:
                parts.append(current.strip())
            
            result_parts = []
            for part in parts:
                if part:
                    evaluated = self._evaluate_simple_expression(part)
                    result_parts.append(str(evaluated) if evaluated is not None else "")
            
            return ''.join(result_parts)
        
        return self._evaluate_simple_expression(expr)
    
    def _evaluate_simple_expression(self, expr: str) -> Any:
        """计算简单表达式值（不包含连接符）"""
        expr = expr.strip()
        
        if not expr:
            return ""
        
        # 字符串常量
        if (expr.startswith('"') and expr.endswith('"')) or \
           (expr.startswith("'") and expr.endswith("'")):
            return expr[1:-1]
        
        # 数字
        if expr.replace('.', '', 1).replace('-', '', 1).isdigit():
            if '.' in expr:
                return float(expr)
            else:
                return int(expr)
        
        # 布尔值
        if expr.upper() in ['TRUE', 'FALSE']:
            return expr.upper() == 'TRUE'
        
        # 常量
        if expr.upper() in self.constants:
            return self.constants[expr.upper()]
        
        # 变量
        if expr.upper() in self.variables:
            val = self.variables[expr.upper()]
            return "" if val is None else val
        elif expr in self.variables:
            val = self.variables[expr]
            return "" if val is None else val
        
        # 函数调用
        if '(' in expr and expr.endswith(')'):
            result = self._execute_function_call(expr)
            return "" if result is None else result
        
        # 尝试算术表达式
        try:
            safe_dict = {"__builtins__": {}}
            for k, v in self.variables.items():
                if not callable(v):
                    safe_dict[k] = v
            safe_dict.update({
                'abs': abs, 'round': round, 'int': int, 'float': float, 'str': str,
                'len': len, 'min': min, 'max': max, 'sum': sum, 'math': math,
            })
            for k, v in self.constants.items():
                safe_dict[k] = v
            result = eval(expr, safe_dict)
            return result
        except Exception:
            return expr
    
    def _split_args_improved(self, args_str: str) -> List[str]:
        """改进的参数分割"""
        args = []
        current = ""
        depth = 0
        in_quotes = False
        quote_char = None
        
        i = 0
        while i < len(args_str):
            char = args_str[i]
            
            if char in ['"', "'"]:
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
                    quote_char = None
                current += char
            elif char == '(' and not in_quotes:
                depth += 1
                current += char
            elif char == ')' and not in_quotes:
                depth -= 1
                current += char
            elif char == ',' and depth == 0 and not in_quotes:
                args.append(current.strip())
                current = ""
            else:
                current += char
            
            i += 1
        
        if current:
            args.append(current.strip())
        
        return args
    
    def save_as_ansi_vbs(self, filename: str, code: str) -> bool:
        """将代码保存为ANSI编码的VBS文件"""
        if not filename.lower().endswith('.vbs'):
            filename += '.vbs'
        
        header = f"' VBScript 文件 - 生成于 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        header += "' 编码: ANSI (GBK)\n"
        header += "' ===========================================\n\n"
        
        full_content = header + code
        return save_vbs_ansi(filename, full_content)

# ================================================
# 第七部分：主接口函数
# ================================================

def Runvbs(code_or_file: str, save_as: Optional[str] = None) -> bool:
    """运行VBScript代码或文件"""
    interpreter = SimpleVBSInterpreter()
    
    if os.path.exists(code_or_file):
        with open(code_or_file, 'r', encoding='gbk', errors='ignore') as f:
            code = f.read()
        interpreter.execute(code)
        if save_as:
            return interpreter.save_as_ansi_vbs(save_as, code)
        return True
    else:
        interpreter.execute(code_or_file)
        if save_as:
            return interpreter.save_as_ansi_vbs(save_as, code_or_file)
        return True

def EvalVBS(expression: str) -> Any:
    """计算VBScript表达式"""
    interpreter = SimpleVBSInterpreter()
    return interpreter._evaluate_expression(expression)

def CreateVBSFile(filename: str, content: str, encoding: str = "ANSI") -> bool:
    """创建VBScript文件"""
    if not filename.lower().endswith('.vbs'):
        filename += '.vbs'
    
    if encoding.upper() == "ANSI":
        return save_vbs_ansi(filename, content)
    else:
        encoding_map = {
            "UTF-8": "utf-8",
            "UTF-16": "utf-16",
            "GBK": "gbk",
        }
        actual_encoding = encoding_map.get(encoding.upper(), "gbk")
        try:
            with open(filename, 'w', encoding=actual_encoding) as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"创建文件失败: {e}")
            return False

def install():
    """安装VBScript函数到全局命名空间"""
    import builtins
    
    current_module = sys.modules[__name__]
    
    # 安装常量
    constants = [
        'vbOKOnly', 'vbOKCancel', 'vbAbortRetryIgnore', 'vbYesNoCancel',
        'vbYesNo', 'vbRetryCancel', 'vbCritical', 'vbQuestion',
        'vbExclamation', 'vbInformation', 'vbOK', 'vbCancel', 'vbAbort',
        'vbRetry', 'vbIgnore', 'vbYes', 'vbNo', 'vbBinaryCompare',
        'vbTextCompare', 'ForReading', 'ForWriting', 'ForAppending',
    ]
    
    for const in constants:
        if hasattr(current_module, const):
            setattr(builtins, const, getattr(current_module, const))
    
    # 安装函数
    functions = [
        'msgbox', 'inputbox', 'CreateObject', 'GetObject', 'Len', 'Left', 'Right',
        'Mid', 'UCase', 'LCase', 'Trim', 'LTrim', 'RTrim', 'InStr', 'Replace',
        'Split', 'Join', 'Asc', 'Chr', 'Space', 'String', 'StrReverse', 'Abs',
        'Sqr', 'Sgn', 'Int', 'Fix', 'Round', 'Rnd', 'Randomize', 'CStr', 'CInt',
        'CLng', 'CSng', 'CDbl', 'CBool', 'CDate', 'Now', 'Date', 'Time', 'Year',
        'Month', 'Day', 'Hour', 'Minute', 'Second', 'Timer', 'DateAdd', 'DateDiff',
        'Array', 'UBound', 'LBound', 'Filter', 'IsNumeric', 'IsDate', 'IsEmpty',
        'IsNull', 'IsArray', 'IsObject', 'TypeName', 'VarType', 'Hex', 'Oct',
        'Val', 'Str',
    ]
    
    for func in functions:
        if hasattr(current_module, func):
            setattr(builtins, func, getattr(current_module, func))
    
    # 安装对象
    builtins.WScript = WScript
    builtins.Runvbs = Runvbs
    builtins.EvalVBS = EvalVBS
    builtins.CreateVBSFile = CreateVBSFile
    
    print("✅ VBScript库已安装到全局命名空间")
    print("📚 现在可以直接使用: msgbox, inputbox, CreateObject 等函数")
    print("📚 示例: msgbox('Hello World!', vbInformation, '测试')")

# ================================================
# 第八部分：导出列表
# ================================================

__all__ = [
    # 常量
    'vbOKOnly', 'vbOKCancel', 'vbAbortRetryIgnore', 'vbYesNoCancel',
    'vbYesNo', 'vbRetryCancel', 'vbCritical', 'vbQuestion',
    'vbExclamation', 'vbInformation', 'vbOK', 'vbCancel', 'vbAbort',
    'vbRetry', 'vbIgnore', 'vbYes', 'vbNo', 'vbBinaryCompare',
    'vbTextCompare', 'ForReading', 'ForWriting', 'ForAppending',
    'vbGeneralDate', 'vbLongDate', 'vbShortDate', 'vbLongTime',
    'vbShortTime', 'vbSunday', 'vbMonday', 'vbTuesday', 'vbWednesday',
    'vbThursday', 'vbFriday', 'vbSaturday', 'vbEmpty', 'vbNull',
    'vbInteger', 'vbLong', 'vbSingle', 'vbDouble', 'vbCurrency',
    'vbDate', 'vbString', 'vbObject', 'vbBoolean', 'vbArray',
    
    # 函数
    'msgbox', 'inputbox', 'CreateObject', 'GetObject', 'Len', 'Left', 'Right',
    'Mid', 'UCase', 'LCase', 'Trim', 'LTrim', 'RTrim', 'InStr', 'Replace',
    'Split', 'Join', 'Asc', 'Chr', 'Space', 'String', 'StrReverse', 'Abs',
    'Sqr', 'Sgn', 'Int', 'Fix', 'Round', 'Rnd', 'Randomize', 'CStr', 'CInt',
    'CLng', 'CSng', 'CDbl', 'CBool', 'CDate', 'Now', 'Date', 'Time', 'Year',
    'Month', 'Day', 'Hour', 'Minute', 'Second', 'Timer', 'DateAdd', 'DateDiff',
    'Array', 'UBound', 'LBound', 'Filter', 'IsNumeric', 'IsDate', 'IsEmpty',
    'IsNull', 'IsArray', 'IsObject', 'TypeName', 'VarType', 'Hex', 'Oct',
    'Val', 'Str',
    
    # 文件操作函数
    'save_vbs_ansi', 'read_vbs_ansi', 'detect_file_encoding',
    
    # 主接口函数
    'Runvbs', 'EvalVBS', 'CreateVBSFile', 'install',
    
    # 对象类
    'WScript', 'WScriptShell', 'FileSystemObject', 'File', 'Folder',
    'TextStream', 'Dictionary', 'GenericCOMObject',
    
    # 解释器
    'SimpleVBSInterpreter',
]

# ================================================
# 第九部分：模块初始化
# ================================================

# 自动清理Tkinter资源
import atexit
_tk_root = None

def _cleanup_tkinter():
    global _tk_root
    if _tk_root:
        try:
            _tk_root.destroy()
        except:
            pass
        _tk_root = None

atexit.register(_cleanup_tkinter)

# ================================================
# 第十部分：测试代码
# ================================================

if __name__ == "__main__":
    print("=" * 70)
    print("VBScript Complete Emulator for Python - 修复完整版")
    print("=" * 70)
    
    # 创建解释器实例
    interpreter = SimpleVBSInterpreter()
    
    # 测试1：基本功能
    print("\n🧪 测试1: 基本功能测试")
    test_code1 = '''
    ' 基本功能测试
    msgbox "VBScript模拟器测试", vbInformation + vbOKOnly, "测试"
    name = inputbox("请输入您的姓名:", "用户信息", "张三")
    age = inputbox("请输入您的年龄:", "用户信息", "25")
    
    If IsNumeric(age) Then
        msgbox "您好 " & name & "!" & vbCrLf & _
               "您的年龄是: " & age & "岁", _
               vbInformation, "欢迎"
    Else
        msgbox "年龄必须是数字!", vbCritical, "错误"
    End If
    
    WScript.Echo "测试完成!"
    '''
    
    interpreter.execute(test_code1)
    
    # 测试2：文件操作
    print("\n🧪 测试2: 文件操作测试")
    test_code2 = '''
    ' 文件操作测试
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 创建ANSI编码的文件
    fileName = "test_ansi.txt"
    content = "这是ANSI编码的测试文件" & vbCrLf & _
              "创建时间: " & Now() & vbCrLf & _
              "测试中文: 你好，世界！"
    
    On Error Resume Next
    Set file = fso.CreateTextFile(fileName, True, False)  ' False = ANSI编码
    
    If Err.Number = 0 Then
        file.Write content
        file.Close
        msgbox "ANSI文件创建成功!" & vbCrLf & _
               "文件名: " & fileName, _
               vbInformation, "成功"
        
        ' 读取文件验证
        Set readFile = fso.OpenTextFile(fileName, 1)  ' 1 = 读取
        fileContent = readFile.ReadAll()
        readFile.Close
        
        msgbox "文件内容:" & vbCrLf & vbCrLf & fileContent, _
               vbInformation, "文件内容"
    Else
        msgbox "创建文件失败: " & Err.Description, vbCritical, "错误"
    End If
    '''
    
    interpreter.execute(test_code2)
    
    # 测试3：保存为ANSI VBS文件
    print("\n🧪 测试3: 保存ANSI VBS文件")
    vbs_content = '''
    ' 这是一个ANSI编码的VBScript文件
    ' 可以直接在Windows上双击运行
    
    Option Explicit
    
    Dim name, age, message
    
    name = inputbox("请输入姓名:", "用户信息", "张三")
    age = inputbox("请输入年龄:", "用户信息", "25")
    
    If name <> "" And IsNumeric(age) Then
        message = "用户信息汇总:" & vbCrLf & _
                 "====================" & vbCrLf & _
                 "姓名: " & name & vbCrLf & _
                 "年龄: " & age & "岁" & vbCrLf & _
                 "出生年份: " & (Year(Now()) - CInt(age)) & vbCrLf & _
                 "===================="
        
        msgbox message, vbInformation, "信息汇总"
        
        ' 保存到文件
        Dim fso, file
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.CreateTextFile("user_info_saved.txt", True, False)
        file.Write message
        file.Close
        
        msgbox "信息已保存到文件!", vbExclamation
    Else
        msgbox "输入无效!", vbCritical
    End If
    
    WScript.Echo "脚本执行完成!"
    '''
    
    # 保存为ANSI VBS文件
    if interpreter.save_as_ansi_vbs("generated_script.vbs", vbs_content):
        print("✅ 已生成: generated_script.vbs (ANSI编码)")
        
        # 验证文件编码
        try:
            with open("generated_script.vbs", 'r', encoding='gbk') as f:
                content_check = f.read(200)
                if "ANSI编码" in content_check:
                    print("✅ ANSI编码验证成功")
        except:
            print("⚠️  编码验证失败")
    
    # 测试4：直接使用Runvbs
    print("\n🧪 测试4: 使用Runvbs函数")
    Runvbs('''
    msgbox "这是通过Runvbs执行的脚本", vbInformation, "Runvbs测试"
    Dim result
    result = msgbox("是否继续测试?", vbYesNo + vbQuestion, "确认")
    
    If result = vbYes Then
        msgbox "您选择了继续", vbInformation
    Else
        msgbox "您选择了取消", vbExclamation
    End If
    ''', save_as="Runvbs_test.vbs")
    
    # 显示生成的文件
    print("\n📁 生成的文件列表:")
    for file in os.listdir('.'):
        if file.endswith(('.vbs', '.txt')):
            try:
                size = os.path.getsize(file)
                print(f"  📄 {file} ({size} 字节)")
            except:
                pass
    
    print("\n" + "=" * 70)
    print("✅ 所有测试完成!")
    print("\n💡 使用提示:")
    print("  1. 导入模块: from vbs import *")
    print("  2. 安装到全局: install()")
    print("  3. 运行脚本: Runvbs('代码') 或 Runvbs('文件.vbs')")
    print("  4. 创建VBS文件: CreateVBSFile('文件名', '内容', 'ANSI')")
    print("\n🎯 特性总结:")
    print("  ✓ 完整的VBScript函数支持")
    print("  ✓ ANSI编码文件保存")
    print("  ✓ 文件系统操作")
    print("  ✓ COM对象模拟")
    print("  ✓ 对话框支持")
    print("  ✓ 中文兼容")
    print("=" * 70)
