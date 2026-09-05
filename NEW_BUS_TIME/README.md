# NEW_BUS_TIME

אפליקציית ווב פשוטה (קובץ `index.html` יחיד) להצגת אתר **busnearby** לפי תחנת אוטובוס נבחרת.

- **כתובת חיה:** https://roobidv.github.io/windsurf-project/NEW_BUS_TIME/
- **מאגר Git:** https://github.com/roobidv/windsurf-project (ענף `main`, GitHub Pages)
- **מיקום מקומי:** `C:\Users\USER\Dropbox\VB6\VBA\CascadeProjects\windsurf-project\NEW_BUS_TIME\`

## מבנה

| קובץ/תיקייה | תפקיד |
|---|---|
| `index.html` | כל האפליקציה — HTML + CSS + JavaScript בקובץ אחד |
| `images/` | תמונות התחנות (JPG, מכווצות) |

## רשימת התחנות (נכון ל-05/09/2026)

### קבוצה ראשית (`STOPS`)

| # | קוד | שם | קטגוריה | צבע כפתור | תמונה | קישור busnearby |
|---|---|---|---|---|---|---|
| 1 | 37018 | מהשכונה טשרנחובסקי ויצמן | base | ירוק | 37018.jpg | 1:24394 |
| 2 | 22942 | מחלף גבעת שמואל/כביש 4 | base | ירוק | 22942.jpg | 1:13274 |
| 3 | 33737 | צומת בית דגן | base | ירוק | 33737.jpg | 1:21799 |
| 4 | 32465 | מליד העבודה של אבא | base | ירוק | 32465.jpg | 1:16390 |
| 5 | 33442 | מהבסיס תחנה שחורה | home | אדום | 33442.jpg | 1:21596 |
| 6 | 33440 | צומת בית דגן | home | אדום | 33440.jpg | 1:21595 |
| 7 | 16464 | בהדים רחוק | home | **כתום** | 16464.jpg | 1:28549 |
| 8 | 38243 | ליד עבודה אבא לבית | home | אדום | 38243.jpg | 1:25206 |
| 9 | 27063 | צומת גהה הביתה | home | אדום | 27063.jpg | 1:14348 |

### תחנות נוספות (`EXTRAS` — כחול)

| # | קוד | שם | קישור |
|---|---|---|---|
| 10 | 20349 | סבידור/נמיר | 1:29480 |
| 11 | 21644 | אבא הלל/זבוט | 1:13198 |

### עוד (`MORE` — כתום, רוחב 48%)

| # | שם | קישור |
|---|---|---|
| 12 | ב"ש מרכז | 1:26652 |
| 13 | בהדים | 1:38736 |

## איך עובד הקוד

- שלושה מערכים בראש ה-`<script>`: `STOPS`, `EXTRAS`, `MORE`.
- מספר הכפתור נקבע אוטומטית לפי הסדר (`btnNum`).
- `dir` קובע צבע הכרטיס: `base`=ירוק, `home`=אדום, `extra`=כחול.
- `btnClass` (אופציונלי) דורס את צבע הכפתור:
  - `btn-orange` — כתום ברוחב 48% (לזוגות)
  - `btn-orange btn-fill` — כתום ברוחב מלא כמו שאר הכפתורים (כפתור 7)
- `img` — שם קובץ התמונה בתוך `images/` (ריק = בלי תמונה).
- `url` — קישור ל-busnearby בפורמט `https://busnearby.co.il/stop/1:<GTFS stop_id>`.

## הוספת/עריכת תחנה

1. מוצאים את מזהה ה-GTFS של התחנה (למשל דרך OpenStreetMap — `gtfs:stop_id:IL-MOT`).
2. שמים תמונה מכווצת ב-`images/` (מומלץ ~600px רוחב, איכות JPEG 60, עד ~20KB).
3. מוסיפים שורה למערך המתאים ב-`index.html`.
4. מבצעים commit + push ל-`main` — האתר מתעדכן תוך דקה-שתיים.

## כיווץ תמונה (PowerShell)

```powershell
Add-Type -AssemblyName System.Drawing
$src=[System.Drawing.Image]::FromFile("SOURCE.jpg")
$bmp=New-Object System.Drawing.Bitmap(600,[int]($src.Height*600/$src.Width))
$g=[System.Drawing.Graphics]::FromImage($bmp)
$g.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.DrawImage($src,0,0,600,$bmp.Height)
$codec=[System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ? {$_.MimeType -eq "image/jpeg"}
$ep=New-Object System.Drawing.Imaging.EncoderParameters(1)
$ep.Param[0]=New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,[long]60)
$bmp.Save("images\XXXXX.jpg",$codec,$ep)
$g.Dispose(); $bmp.Dispose(); $src.Dispose()
```

## היסטוריית שינויים אחרונים

| תאריך | commit | תיאור |
|---|---|---|
| 05/09/2026 | `7196323` | הוספת תמונה מכווצת לתחנה 16464 |
| 05/09/2026 | `bc22cb6` | כפתור 7: שונה ל-"בהדים רחוק" (16464), כתום, קישור 1:28549 |
| 05/09/2026 | `e147173` | כפתור 7: רוחב מלא (btn-fill) במקום 48% |
