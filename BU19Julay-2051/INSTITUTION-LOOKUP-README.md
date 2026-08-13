# כלי חיפוש מוסדות חינוך בישראל

## סקירה כללית

כלי אינטרנטי (HTML5) לחיפוש מוסדות חינוך בישראל לפי סמל מוסד (6 ספרות) או לפי שם מוסד + עיר.
הכלי פועל כדף HTML עצמאי ללא צורך בשרת — מתאים להרצה מקומית או דרך GitHub Pages.

**כתובת פעילה:** https://roobidv.github.io/windsurf-project/institution-lookup.html

**מחבר:** רובי דביר (roobid@gmail.com)

---

## קבצי הפרויקט

| קובץ | תיאור | שורות |
|---|---|---|
| `institution-lookup.html` | דף HTML ראשי — ממשק משתמש, עיצוב, לוגיקה | 868 |
| `institution_lookup_gas.js` | סקריפט Google Apps Script — proxy לנתונים מורחבים | 234 |

---

## ארכיטקטורה

```
┌─────────────────────────────────┐
│   institution-lookup.html       │
│   (דפדפן המשתמש)               │
│                                 │
│   חיפוש לפי סמל ← קריאת API    │
│   חיפוש לפי שם+עיר ← קריאת API │
└──────────┬──────────┬───────────┘
           │          │
     ┌─────▼──────┐ ┌─▼──────────────────┐
     │ data.gov.il│ │ Google Apps Script  │
     │ (API ישיר) │ │ (Proxy)            │
     │            │ │                    │
     │ נתונים     │ │  ┌──────────────┐  │
     │ בסיסיים:   │ │  │allschool.co.il│ │
     │ - שם       │ │  │ טלפון, כתובת │  │
     │ - ישוב     │ │  │ מנהל/ת       │  │
     │ - מגזר     │ │  └──────────────┘  │
     │ - פיקוח    │ │                    │
     │ - שכבות    │ │  ┌──────────────┐  │
     │ - תלמידים  │ │  │serviced.co.il│  │
     │ - שנת יסוד │ │  │ דוא"ל בי"ס   │  │
     └────────────┘ │  └──────────────┘  │
                    └────────────────────┘
```

### מקור 1: data.gov.il (API ישיר, ללא CORS)

- **URL:** `https://data.gov.il/api/3/action/datastore_search`
- **Resource ID:** `5548fd63-5868-4053-ad81-98caddc5e232`
- **חינם, ללא הרשמה, ללא מפתח API**
- **שדות:** שם מוסד, ישוב, רשות, מגזר, פיקוח, שכבות (מ/עד), סוג מסגרת, תלמידים, מעמד משפטי, שנת יסוד, מחוז

### מקור 2: Google Apps Script (Proxy)

- **URL נוכחי:** `https://script.google.com/macros/s/AKfycbwq-h7h84PAPMwJ6r_ggKgfGnT_xnZZSRBwNO1823FlD1nFXf7WmnQ6on9LPN-8NNZh/exec`
- **מגרד (scrapes) מ:**
  - **allschool.co.il** — טלפון מזכירות, כתובת מלאה (רחוב + מספר), מנהל/ת
  - **serviced.co.il** — דוא"ל בית ספר (בד"כ @hinuchm.k12.il)
- **שימוש ישיר:**
  ```
  GET ...exec?semel=310011
  → { semel, name, address, phone, principal, email, source }
  ```

---

## תכונות

### חיפוש לפי סמל מוסד
הזנת 6 ספרות → חיפוש ישיר ב-API + GAS → הצגת כל הנתונים

### חיפוש לפי שם + עיר
הזנת טקסט (לא ספרות בלבד) → חיפוש ב-data.gov.il עם פילטר לפי עיר → הצגת רשימת תוצאות → לחיצה על תוצאה מציגה פרטים מלאים

### כרטיס מוסד
ריבוע ירוק עם סיכום: סמל, שם, טלפון, כתובת, דוא"ל + כפתור "העתק ללוח גזירים"

### העתקה ללוח גזירים
לחיצה אחת מעתיקה טקסט מסודר:
```
סמל מוסד: 310011
שם מוסד: גימס רוטשילד, אור עקיבא
טלפון: 04-6361277
כתובת: שד ירושלים, אור עקיבא
דוא"ל: meiravtur@hinuchm.k12.il
```

### תצוגה מותאמת לסלולרי
פונטים מוקטנים, כפתורים קומפקטיים, input ו-button בשורה אחת גם ב-360px

---

## התקנה והפעלה

### אפשרות 1: GitHub Pages (מומלץ)
1. ב-GitHub → Settings → Pages → Source: Deploy from branch → Branch: `main` → Save
2. הדף זמין ב: `https://roobidv.github.io/windsurf-project/institution-lookup.html`

### אפשרות 2: הרצה מקומית
1. הורד את `institution-lookup.html`
2. פתח בדפדפן — עובד בלי שרת

### אפשרות 3: שרת מקומי
```bash
python3 -m http.server 8080
# פתח http://localhost:8080/institution-lookup.html
```

---

## הגדרת Google Apps Script (אופציונלי — לנתונים מורחבים)

ללא GAS, הדף מציג רק נתונים בסיסיים מ-data.gov.il.
עם GAS, מתווספים: טלפון, כתובת מלאה, מנהל/ת, דוא"ל.

### הגדרה:
1. לך ל-https://script.google.com → פרויקט חדש
2. מחק את הקוד הקיים, הדבק את התוכן מ-`institution_lookup_gas.js`
3. Deploy → New deployment → Type: Web app → Who has access: Anyone → Deploy
4. העתק את ה-URL שמתקבל

### עדכון Deploy קיים:
1. ב-https://script.google.com → פתח את הפרויקט
2. Deploy → Manage deployments → Edit (אייקון עיפרון) → Version: New version → Deploy

**ה-URL הנוכחי כבר מוטמע בקוד ה-HTML כברירת מחדל.**

---

## היסטוריית גרסאות (PRs)

| PR | תיאור |
|---|---|
| #13 | יצירת הכלי — חיפוש לפי סמל, ממשק עברי RTL, חיבור ל-API |
| #14 | תיקון תצוגת שכבות — "גן" במקום "0" |
| #15 | כרטיס מוסד ירוק + כפתור העתקה ללוח גזירים |
| #16 | דוא"ל במקום מפקח/ת, שורת מחבר צהובה, הסתרת מקורות מידע, דיסקליימר |
| #17 | הסרת שורת "נתונים ממשרד החינוך" |
| #18 | חיפוש לפי שם מוסד + עיר (חלקי) |
| #19 | הסרת שורת נתונים (חוזר) |
| #20 | חיפוש דוא"ל ב-serviced.co.il כ-fallback |
| #21 | עדכון URL של GAS ל-Deploy חדש |
| #22 | כפתור חיפוש קטן יותר, הסרת שורת footer, תיקון קידוד גרשיים |
| #23 | תיקון תצוגת סלולרי — מניעת גלישה, הקטנת פונטים |
| #24 | תיקון קידוד HTML entities בשמות מוסדות (חט"ב) |
| #25 | תיקון כפתור העתקה עם שמות שמכילים גרשיים |

---

## מבנה טכני

### HTML (institution-lookup.html)
- **שורות 1-460:** HTML + CSS (עיצוב, responsive, RTL)
- **שורות 460-868:** JavaScript (API calls, rendering, copy, search logic)
- **Self-contained:** אין תלויות חיצוניות (ללא jQuery, Bootstrap וכו')
- **קידוד:** UTF-8, עם פונקציות `cleanQuotes()` ו-`decodeHtml()` לטיפול בתווים מיוחדים

### GAS Script (institution_lookup_gas.js)
- **פונקציה ראשית:** `doGet(e)` — מקבל `semel` כפרמטר, מחזיר JSON
- **scrapeAllSchool(semel):** גורד allschool.co.il → טלפון, כתובת, מנהל/ת
- **scrapeServicedEmail(schoolName):** חיפוש + גריד serviced.co.il → דוא"ל
- **מטפל ב-CORS** — הדפדפן לא יכול לגשת ישירות לאתרים אלו

---

## הערות חשובות

- **הנתונים מ-data.gov.il הם משנת 2015** (הגרסה האחרונה הזמינה ב-API)
- **allschool.co.il ו-serviced.co.il** — אתרים חיצוניים שעלולים להשתנות. אם ה-scraping נשבר, הדף ימשיך לעבוד עם נתונים בסיסיים
- **אין שימוש מסחרי** — הנתונים נלקחים ממקורות חופשיים וגלויים, אך יש איסור על שימוש מסחרי
- **localStorage** — ה-URL של GAS נשמר ב-localStorage של הדפדפן (אם הוזן ידנית)
