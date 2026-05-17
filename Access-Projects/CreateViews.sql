-- ============================================================================
-- סקריפט SQL ליצירת תצוגות חייגן טלפון ב-Access - חלק 2: תצוגות
-- Author: VBA Developer
-- Date: 31/03/2026
-- הפעל כל פקודה בנפרד!
-- ============================================================================

-- שאילתה 1: אנשי קשר פעילים
CREATE VIEW ActiveContacts AS
SELECT ContactID, ContactName, PhoneNumber, Email, DateAdded, CallCount
FROM Contacts
ORDER BY CallCount DESC, ContactName;

-- שאילתה 2: היסטוריית שיחות אחרונה
CREATE VIEW RecentCalls AS
SELECT c.ContactName, h.PhoneNumber, h.CallDate, h.CallTime, h.CallDuration, h.CallType
FROM CallHistory h
LEFT JOIN Contacts c ON h.PhoneNumber = c.PhoneNumber
WHERE h.CallDate >= Date() - 30
ORDER BY h.CallDate DESC, h.CallTime DESC;

-- שאילתה 3: אנשי קשר שלא דיברתי איתם הרבה זמן
CREATE VIEW LongTimeNoContact AS
SELECT ContactName, PhoneNumber, CallCount
FROM Contacts
ORDER BY CallCount DESC, ContactName;
