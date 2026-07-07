# 🏥 نظام البريكات - Break Schedule
### مستشفيات المانع العامة - Al Mana General Hospitals

نظام إدارة بريكات الموظفين مع مزامنة تلقائية بين الإدارة والموظفين.

## 🔗 الرابطين

| الرابط | الاستخدام |
|-------|-----------|
| `schedule.html` | 👨‍💼 الإدارة (عرض + بريكاتي + إدارة) |
| `schedule.html?mode=emp` | 👤 الموظفين (اختيار بريكات + تذكير) |

## 🚀 النشر على PythonAnywhere (مجاني)

### 1️⃣ رفع الملفات على GitHub
```
1. https://github.com/new → اسم: break-schedule → Public
2. ارفع الملفات: schedule.html, app.py, requirements.txt, Break_Schedule_Enhanced.xlsm
```

### 2️⃣ إنشاء حساب PythonAnywhere
```
1. https://pythonanywhere.com → Register (مجاني)
2. Dashboard → Web → Add a new web app
3. Next → Manual configuration → Python 3.10 → Next
```

### 3️⃣ رفع الملفات من GitHub
```bash
# في PythonAnywhere Bash Console:
git clone https://github.com/YourUsername/break-schedule.git
cd break-schedule
pip install -r requirements.txt --user
```

أو ارفع الملفات يدوي من Files tab.

### 4️⃣ تعديل WSGI
```
في Web tab:
1. Code → WSGI configuration file
2. احذف كل شيء واكتب:

import sys
sys.path.insert(0, '/home/YourUsername/break-schedule')
from app import app as application

3. Save
```

### 5️⃣ تشغيل
```
في Web tab → Reload → ✅
```

### 6️⃣ روابطك
```
https://YourUsername.pythonanywhere.com/schedule.html      ← الإدارة
https://YourUsername.pythonanywhere.com/schedule.html?mode=emp  ← الموظفين
```

## 💻 التشغيل المحلي
```bash
pip install -r requirements.txt
python3 app.py
```
ثم افتح:
- http://localhost:5000/schedule.html
- http://localhost:5000/schedule.html?mode=emp

## 📁 الملفات
- `schedule.html` — الواجهة (إدارة + موظفين في ملف واحد)
- `app.py` — سيرفر Flask (API + Excel + إيميل + قاعدة بيانات SQLite)
- `requirements.txt` — المكتبات المطلوبة
- `Break_Schedule_Enhanced.xlsm` — قالب الإكسل

## المميزات
- ✅ ملف واحد — رابطين (إدارة + موظف)
- ☁️ مزامنة تلقائية عبر السيرفر (SQLite)
- 🔊 جرس صوتي + 🔔 إشعار متصفح
- 📧 إرسال إيميل عبر SMTP
- 📥 تصدير إكسل
- 🕐 إضافة/حذف/تعديل أوقات ومدد
- 🔄 تصفير تلقائي كل 24 ساعة
