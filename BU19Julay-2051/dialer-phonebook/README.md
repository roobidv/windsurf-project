# Dialer PhoneBook - ספר טלפונים

אפליקציית ספר טלפונים מבוססת GitHub Pages, עם סנכרון ל-Google Sheets.

## מבנה

```
dialer-phonebook/
├── index.html      ← האפליקציה (קובץ יחיד)
├── code.gs         ← Google Apps Script (להדביק ב-Sheets)
└── README.md       ← קובץ זה
```

## התקנה

### 1. Google Apps Script

1. פתח את קובץ **Dialer Contacts** ב-Google Sheets
2. Extensions → Apps Script
3. מחק את כל התוכן הקיים ב-`Code.gs`
4. הדבק את תוכן הקובץ `code.gs` מהפרויקט
5. Deploy → New deployment → Web app
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the URL

### 2. הגדרת URL באפליקציה

1. פתח `index.html`
2. מצא: `const SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';`
3. החלף ב-URL שקיבלת מה-Deploy

### 3. GitHub Pages

1. צור repository חדש ב-GitHub (למשל `phonebook`)
2. העלה את `index.html` ל-root
3. Settings → Pages → Source: main branch
4. האתר יהיה זמין ב: `https://USERNAME.github.io/phonebook/`

## גליונות נדרשים ב-Dialer Contacts

| גליון | תפקיד |
|--------|--------|
| Contacts | נתוני אנשי קשר |
| הרשאות | אימות משתמשים (טלפון/מייל/קוד/שם/הרשאה) |
| מכשירים | רישום מכשירים מאומתים |
| ארועים | לוג פעולות |

## עמודות גליון Contacts

ContactID, ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Address, Notes, DateAdded, CallCount, CUS_NUMBER

## שימוש

- **חיפוש** — שם, טלפון, מייל, הערות, מספר לקוח
- **לחיצה על כרטיס** — פרטים מורחבים
- **📱 מספר טלפון** — לחיצה = חיוג
- **WhatsApp** — כפתור ירוק (רק למספרי 05x)
- **📋 העתק** — העתק מספר ללוח
- **📧 מייל** — פתיחת אפליקציית מייל

## אבטחה

- כניסה דרך טלפון/מייל/קוד מול גליון הרשאות
- מכשיר מזוהה נכנס אוטומטית בפעמים הבאות
- כל כניסה נרשמת בגליון ארועים
