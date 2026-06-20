# תיעוד מלא - סביבת הפיתוח

> עדכון אחרון: יוני 2026

---

## סקירה כללית

פרויקט זה מכיל שלושה מוצרים עיקריים:

| מוצר | תיאור | טכנולוגיה |
|------|--------|-----------|
| **Logisty2026** | מערכת ניהול לקוחות ומכירות | HTML + Google Apps Script |
| **Dialer PhoneBook** | ספר טלפונים עם חיוג, WhatsApp, Dropbox | HTML + Google Apps Script |
| **Access Dialer** | חייגן טלפון שולחני | VBA + Microsoft Access |

---

## 1. Logisty2026 - מערכת ניהול לקוחות

### קבצים
| קובץ | תפקיד |
|-------|--------|
| `logisty2026.html` | הגרסה המקוונת (עם Google Sheets) |
| `logisty2026_offline.html` | הגרסה לשימוש offline (cache מקומי) |
| `Code_gs.js` | הסקריפט של Google Apps Script |

### כתובת Deployment
```
https://script.google.com/macros/s/AKfycbyn_zSxAa3salEw9PT9voZT9ULWnHBUdSVUAP8kIC5mIIZg8ic6PLo4wkIhBAblEiZbOw/exec
```

### פונקציות Backend (Code_gs.js)
| Action | Method | תיאור |
|--------|--------|--------|
| `getCSV` | GET | שליפת כל הנתונים כ-CSV |
| `auth` | GET | אימות משתמש (phone/email/code) |
| `getUserEmail` | GET | שליפת מייל משתמש |
| `getReport` | GET | דוח פעילות משתמש |
| `getMgrReport` | GET | דוח מנהל (סיכום לפי איש מכירות) |
| `getContacts` | GET | שליפת כתובות |
| `deviceAuth` | GET | אימות מכשיר |
| `geminiSearch` | GET | חיפוש AI (Gemini) |
| `logLogin` | POST | רישום כניסה |
| `logEvent` | POST | רישום אירוע/שינוי |
| `addNote` | POST | הוספת הערה |
| `registerDevice` | POST | רישום מכשיר |
| `saveField` | POST | שמירת שדה בגיליון |

### גליונות נדרשים (Google Sheets)
| גליון | עמודות עיקריות |
|--------|----------------|
| Sheet1 (ראשי) | סמל, שם מוסד, שלב, סוג רכש, סטטוס, טלפונים, מיילים |
| הרשאות | מזהה, סוג, שם, פעיל, מייל |
| ארועים | תאריך, משתמש, פעולה, שדה, ערך קודם, ערך חדש, סמל |
| הערות | סמל מוסד, שלב, הערה, משתמש, תאריך |
| מכשירים | deviceId, שם, טוקן, UA, תאריך רישום, פעיל, הרשאה |

### הרשאות
- **code** — קוד כניסה (ADMIN2026)
- **phone** — אימות לפי מספר טלפון
- **phone1** — הרשאת טלפון מורחבת
- **email** — אימות לפי מייל
- **domain** — אימות לפי דומיין (@xxx.co.il)

---

## 2. Dialer PhoneBook - ספר טלפונים

### קבצים
| קובץ | תפקיד |
|-------|--------|
| `dialer-phonebook/index.html` | האפליקציה (SPA) |
| `dialer-phonebook/code.gs` | Google Apps Script (Backend) |
| `dialer-phonebook/README.md` | הוראות התקנה |

### כתובת Deployment
```
https://script.google.com/macros/s/AKfycbz0Q1lp-weh9yeBaJXIMXqwo6kNZTLaBMp3TCcckx-MKRlfupKHjO_cAftmG_MLO0VCWQ/exec
```

### פונקציות Backend (code.gs)
| Action | Method | תיאור |
|--------|--------|--------|
| `getCSV` | GET | שליפת אנשי קשר כ-CSV |
| `auth` | GET | אימות (קוד אחיד: 583995) |
| `deviceAuth` | GET | אימות מכשיר |
| `checkMessages` | GET | בדיקת הודעות חדשות |
| `getOfferPdf` | GET | שליפת PDF הצעת מחיר מ-Dropbox |
| `registerDevice` | POST | רישום מכשיר |
| `logLogin` | POST | רישום כניסה |
| `logEvent` | POST | רישום אירוע |
| `updateContact` | POST | עדכון שדה באיש קשר |
| `addNote` | POST | הוספת הערה |
| `markMessageRead` | POST | סימון הודעה כנקראה |

### גליונות נדרשים (Google Sheets - Dialer Contacts)
| גליון | עמודות |
|--------|--------|
| Contacts | ContactID, ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Address, Notes, DateAdded, CallCount, CUS_NUMBER |
| הרשאות | טלפון/מייל/קוד, שם, הרשאה |
| מכשירים | deviceId, name, token, ua, lastSeen, permType |
| ארועים | תאריך, משתמש, פעולה, שדה, ערך קודם, ערך חדש, CUS_NUMBER |
| הודעות | ID, Target, MSG, CreatedDate, ReadDate |

### אבטחה
- קוד כניסה אחיד: `583995`
- הגבלת ניסיונות: 5 ניסיונות ב-5 דקות, חסימה ל-4 שעות
- מכשיר מזוהה נכנס אוטומטית

### אינטגרציית Dropbox
שליפת הצעות מחיר (PDF) מ-Dropbox:

**Script Properties נדרשים:**
| Property | תיאור |
|----------|--------|
| `DROPBOX_REFRESH_TOKEN` | טוקן רענון של Dropbox |
| `DROPBOX_APP_KEY` | מפתח אפליקציית Dropbox |
| `DROPBOX_APP_SECRET` | סוד אפליקציית Dropbox |

**תהליך שליפת קובץ:**
1. קבלת Access Token חדש דרך `oauth2/token` (refresh_token)
2. רשימת קבצים בתיקיית `/Ness/offers` (עם pagination)
3. חיפוש קובץ לפי תבנית: `*_{serial}.pdf`
4. יצירת קישור זמני דרך `get_temporary_link`

**Pagination:** תומך בתיקיות עם יותר מ-100 קבצים (שימוש ב-`list_folder/continue`)

**API Endpoints:**
```
POST https://api.dropbox.com/oauth2/token
POST https://api.dropboxapi.com/2/files/list_folder
POST https://api.dropboxapi.com/2/files/list_folder/continue
POST https://api.dropboxapi.com/2/files/get_temporary_link
```

---

## 3. Access Dialer - חייגן שולחני

### מבנה תיקייה: `Access-Projects/`

### מודולים עיקריים
| קובץ | תפקיד |
|-------|--------|
| `ContactsDialerCode.bas` | קוד ראשי של החייגן |
| `PhoneDialerModule.bas` | לוגיקת חיוג |
| `PhoneDialerForm.frm` | טופס החייגן |
| `GlobalHotkey.bas` | מקשי קיצור (F9) |
| `CloudMsg.bas` | הודעות מהענן |
| `SyncPhoneBook.bas` | סנכרון ספר טלפונים |
| `AutoBackup.bas` | גיבוי אוטומטי |
| `DatabaseUtilities.bas` | כלי DB |
| `ErrorLogger.bas` | רישום שגיאות |
| `FrontEndLinker.bas` | קישור Frontend-Backend |

### טבלאות (CreateTables.sql)
- `tblContacts` — אנשי קשר
- `tblCallHistory` — היסטוריית שיחות
- `tblSettings` — הגדרות מערכת

### דרישות
- Microsoft Access 2016+
- Windows 10/11
- PowerShell (לסקריפט AccessHotkey.ps1)

---

## 4. ספריית VBA משותפת

### תיקייה: `VBA-Library/`
| קובץ | תפקיד |
|-------|--------|
| `StringUtilities.bas` | פונקציות מחרוזת |
| `Template.bas` | תבנית למודול חדש |

### תיקייה: `Excel-Projects/`
| קובץ | תפקיד |
|-------|--------|
| `ExcelUtilities.bas` | כלי עזר ל-Excel |

---

## 5. היסטוריית גרסאות

| גרסה | תאריך | שינויים |
|-------|--------|---------|
| v1.3.0 | יוני 2026 | תיקון הודעות שגיאה (אנגלית בשרת, עברית בלקוח) |
| v1.2.0 | יוני 2026 | מקש F2, הודעות ענן, נתיבים דינמיים |
| - | יוני 2026 | הוספת שם מוסד לדוח פעילות |
| - | מאי 2026 | מניעת תווים לא חוקיים בטלפון |
| - | מאי 2026 | דוח מנהל עם toggle לפירוט |
| - | יוני 2026 | Dropbox pagination (תמיכה ב-100+ קבצים) |

---

## 6. URLs חשובים

| שירות | URL |
|--------|-----|
| GitHub Repo | https://github.com/roobidv/windsurf-project |
| Logisty2026 Script | `AKfycbyn_zSxAa3salEw9PT9voZT9ULWnHBUdSVUAP8kIC5mIIZg8ic6PLo4wkIhBAblEiZbOw` |
| Dialer Script | `AKfycbz0Q1lp-weh9yeBaJXIMXqwo6kNZTLaBMp3TCcckx-MKRlfupKHjO_cAftmG_MLO0VCWQ` |
| Dropbox Folder | `/Ness/offers` |

---

## 7. Script Properties (הגדרות סודיות)

### Dialer PhoneBook Script:
```
DROPBOX_REFRESH_TOKEN = [שמור ב-Script Properties]
DROPBOX_APP_KEY       = [שמור ב-Script Properties]
DROPBOX_APP_SECRET    = [שמור ב-Script Properties]
```

**איך לגשת:** Apps Script → Project Settings (⚙️) → Script Properties

---

## 8. הנחיות Deploy

### עדכון Google Apps Script:
1. Google Sheets → Extensions → Apps Script
2. הדבק/עדכן את הקוד
3. Ctrl+S (שמור)
4. Deploy → Manage deployments → ✏️ → New version → Deploy

### עדכון GitHub Pages (אם משתמשים):
1. `git add .`
2. `git commit -m "description"`
3. `git push origin main`
4. GitHub Pages מתעדכן אוטומטית

---

## 9. Dropbox API - פרמטרים מלאים

### אימות (Token Refresh)
```
POST https://api.dropbox.com/oauth2/token
Params:
  grant_type: 'refresh_token'
  refresh_token: DROPBOX_REFRESH_TOKEN
  client_id: DROPBOX_APP_KEY
  client_secret: DROPBOX_APP_SECRET
Response: { access_token: "..." }
```

### רשימת קבצים (עם Pagination)
```
POST https://api.dropboxapi.com/2/files/list_folder
Headers: Authorization: Bearer {access_token}
Body: { path: '/Ness/offers', limit: 100 }
Response: { entries: [...], has_more: true/false, cursor: "..." }
```

### המשך רשימה (Pagination)
```
POST https://api.dropboxapi.com/2/files/list_folder/continue
Headers: Authorization: Bearer {access_token}
Body: { cursor: "..." }
Response: { entries: [...], has_more: true/false, cursor: "..." }
```

### קישור זמני להורדה
```
POST https://api.dropboxapi.com/2/files/get_temporary_link
Headers: Authorization: Bearer {access_token}
Body: { path: "/ness/offers/filename_1900384.pdf" }
Response: { link: "https://...dl.dropboxusercontent.com/..." }
```

### תנאי חיפוש קובץ
| פרמטר | ערך | הסבר |
|--------|------|-------|
| תיקייה | `/Ness/offers` | תיקייה קבועה |
| תבנית שם | `*_{serial}.pdf` | חיפוש לפי מספר סידורי |
| סוג קובץ | PDF | הצעות מחיר |
| Pagination | `has_more` + `cursor` | ללא הגבלת כמות קבצים |

---

## 10. פרמטרי דואר (Email)

### שליחת מייל עם צרופות (Logisty2026)
```
פרמטרים נשלחים ל-Backend:
  toEmail: כתובת מייל הנמען
  serial: מספר סידורי הצעה
  selectedOffers: רשימת הצעות נבחרות
  user: שם המשתמש השולח
```

### חיפוש בדואר נכנס (Gmail API - לשימוש עתידי)
```
פרמטרי חיפוש:
  q: "has:attachment filename:pdf"
  from: שולח
  subject: נושא
  after/before: טווח תאריכים
  labelIds: [INBOX, SENT]
```
