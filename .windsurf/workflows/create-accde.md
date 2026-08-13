---
description: how to create an ACCDE runtime file from the development ACCDB
---

# נוהל יצירת קובץ ריצה (ACCDE)

## דרישות מוקדמות
- קובץ פיתוח תקין: `C:\Users\USER\Documents\unbound\Database3.accdb`
- כל המודולים מעודכנים (RunUpdate בוצע)
- הטופס הראשי `frmContactsDialer` עובד תקין

## שלבים

### שלב 1 — גיבוי
בצע גיבוי מלא לפני כל שינוי (BAS, FE, BE).

### שלב 2 — הגדרות ידניות ב-Design View
פתח את `frmContactsDialer` ב-Design View ושנה:

| מאפיין | ערך |
|--------|-----|
| **Pop Up** | Yes |
| **Auto Center** | Yes |

> **חשוב:** אל תשנה Border Style, MinMaxButtons, או כל מאפיין עיצוב אחר בקוד.
> כפתור Maximize מוסר אוטומטית ע"י `RemoveMaximizeButton` ב-Form_Load.
> ShortcutMenu (תפריט קליק-ימני) מבוטל אוטומטית ב-Form_Load.
> רוחב הטופס מתרחב ב-5% אוטומטית ב-Form_Load (`AdjustFormWidth`).

### שלב 3 — יצירת ACCDE
1. ב-Access: **File → Save As → Make ACCDE**
2. שמור את הקובץ בתיקייה: `C:\Users\USER\Documents\unbound\קובץ ריצה\Database3.accde`

### שלב 4 — בדיקת ACCDE
פתח את ה-ACCDE וודא:
- [ ] חלון Access מוסתר (רק הטופס נראה)
- [ ] אין כפתור Maximize בפינה הימנית
- [ ] לחיצה ימנית לא מציגה תפריט
- [ ] הטופס ממורכז במסך
- [ ] רוחב הטופס מעט יותר רחב מהמקור
- [ ] טפסי משנה (רשומה חדשה, הגדרות, עריכת שיחה) נפתחים ממורכזים ביחס לטופס הראשי
- [ ] חיוג עובד
- [ ] Extension נרשם ב-CallHistory

### שלב 5 — חזרה לפיתוח
1. **סגור את ה-ACCDE**
2. **החזר את קובץ הפיתוח מהגיבוי** (ללא PopUp=Yes)
3. פתח את ה-ACCDB והמשך בפיתוח רגיל

## מה קורה אוטומטית ב-ACCDE (בקוד)
הקוד ב-`ContactsDialer_Form_Load` מזהה ACCDE ומפעיל:
- `HideAccessFrame` — מזיז את חלון Access מחוץ למסך (MoveWindow API)
- `RemoveMaximizeButton` — מסיר כפתור Maximize (GetWindowLongPtr/SetWindowLongPtr API)
- `AdjustFormWidth` — מרחיב את הטופס ב-5%
- `ShortcutMenu = False` — מבטל תפריט קליק-ימני
- `CenterChildForm` — מרכז טפסי משנה (frmContactEdit, frmSettingsEdit, frmCallHistoryEdit)
- `ShowAccessFrame` — משחזר את חלון Access בסגירה (Form_Unload)

## באגים ידועים לתיקון בגרסה הבאה
- ~~**טפסי משנה נצמדים לפינה השמאלית** — תוקן 02/05/2026 עם GetWindowRect/SetWindowPos API~~

## הערות חשובות
- **לעולם אל תריץ `SetupAccdeProperties`** — זה נועל את קובץ הפיתוח
- **לעולם אל תשנה מאפייני עיצוב בקוד** — רק ידנית ב-Design View
- הקוד מבדיל בין ACCDB ל-ACCDE באמצעות: `LCase(Right$(CurrentDb.Name, 6)) = ".accde"`
