-- ============================================================================
-- סקריפט SQL ליצירת טבלאות חייגן טלפון ב-Access - חלק 1: טבלאות
-- Author: VBA Developer
-- Date: 31/03/2026
-- הפעל כל פקודה בנפרד!
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

-- טבלת אינטראקציות מול אנשי קשר (1-ל-רבים: Contacts -> Interactions)
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
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (1, 'משרד', '031234567', 'מספר טלפון של המשרד');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (2, 'בית', '098765432', 'מספר טלפון ביתי');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (3, 'נייד', '0541234567', 'מספר טלפון נייד');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (4, 'חירום', '100', 'משטרה - חירום');
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description) VALUES (5, 'מידע', '103', 'מידע טלפוני');

-- הוספת אנשי קשר לדוגמה
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('דוד כהן', '0521111111', 'david@email.com', 'עמית לעבודה', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('שרה לוי', '0532222222', 'sara@email.com', 'לקוח חשוב', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('יוסי אברהם', '0543333333', 'yossi@email.com', 'ספק', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('רחל ישראלי', '0504444444', 'rachel@email.com', 'חברה', 0);
INSERT INTO Contacts (ContactName, PhoneNumber, Email, Notes, CallCount) VALUES ('משה דוד', '0555555555', 'moshe@email.com', 'שכן', 0);
