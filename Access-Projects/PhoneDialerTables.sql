-- ============================================================================
-- סקריפט SQL ליצירת טבלאות חייגן טלפון ב-Access
-- Author: VBA Developer
-- Date: 31/03/2026
-- ============================================================================

-- טבלת אנשי קשר
CREATE TABLE Contacts (
    ContactID COUNTER PRIMARY KEY,
    ContactName TEXT(100) NOT NULL,
    PhoneNumber TEXT(20) NOT NULL,
    Landline TEXT(20),
    Email TEXT(100),
    Notes MEMO,
    DateAdded DATE,
    CallCount LONG
);

-- טבלת חיוג מהיר
CREATE TABLE SpeedDial (
    DialID COUNTER PRIMARY KEY,
    DialIndex INTEGER NOT NULL UNIQUE,
    ContactID LONG,
    ContactName TEXT(100) NOT NULL,
    PhoneNumber TEXT(20) NOT NULL,
    Description TEXT(255),
    DateAdded DATE
);

-- טבלת היסטוריית שיחות
CREATE TABLE CallHistory (
    CallID COUNTER PRIMARY KEY,
    ContactID LONG,
    PhoneNumber TEXT(20) NOT NULL,
    ContactName TEXT(100),
    CallDate DATE,
    CallTime TIME,
    CallDuration INTEGER, -- בשניות
    CallType TEXT(20), -- 'Outgoing', 'Incoming', 'Missed'
    Notes MEMO
);

-- טבלת אינטראקציות מול אנשי קשר
CREATE TABLE Interactions (
    InteractionID COUNTER PRIMARY KEY,
    ContactID LONG NOT NULL,
    InteractionDate DATE,
    InteractionTime TIME,
    InteractionType TEXT(30),
    Subject TEXT(255),
    Details MEMO,
    FollowUpDate DATE,
    CreatedAt DATE
);

-- יצירת אינדקסים
CREATE INDEX idx_Contacts_Name ON Contacts(ContactName);
CREATE UNIQUE INDEX idx_Contacts_Phone ON Contacts(PhoneNumber);
CREATE INDEX idx_Contacts_CallCount ON Contacts(CallCount);
CREATE INDEX idx_SpeedDial_ContactID ON SpeedDial(ContactID);
CREATE INDEX idx_CallHistory_Phone ON CallHistory(PhoneNumber);
CREATE INDEX idx_CallHistory_Date ON CallHistory(CallDate, CallTime);
CREATE INDEX idx_CallHistory_ContactID ON CallHistory(ContactID);
CREATE INDEX idx_Interactions_ContactID ON Interactions(ContactID);
CREATE INDEX idx_Interactions_Date ON Interactions(InteractionDate, InteractionTime);

-- הוספת נתוני דוגמה לחיוג מהיר
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (1, 'משרד', '03-1234567', 'מספר טלפון של המשרד');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (2, 'בית', '09-8765432', 'מספר טלפון ביתי');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (3, 'נייד', '054-1234567', 'מספר טלפון נייד');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (4, 'חירום', '100', 'משטרה - חירום');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (5, 'מידע', '103', 'מידע טלפוני');

-- הוספת אנשי קשר לדוגמה
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('דוד כהן', '052-1111111', 'david@email.com', 'עמית לעבודה', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('שרה לוי', '053-2222222', 'sara@email.com', 'לקוח חשוב', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('יוסי אברהם', '054-3333333', 'yossi@email.com', 'ספק', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('רחל ישראלי', '050-4444444', 'rachel@email.com', 'חברה', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('משה דוד', '055-5555555', 'moshe@email.com', 'שכן', 0);

-- יצירת שאילתות שימושיות
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
