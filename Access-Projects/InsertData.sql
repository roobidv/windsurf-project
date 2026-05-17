-- ============================================================================
-- סקריפט SQL להוספת נתונים לחייגן טלפון ב-Access - חלק 3: נתונים
-- Author: VBA Developer
-- Date: 31/03/2026
-- הפעל כל פקודה בנפרד!
-- ============================================================================

-- הוספת נתוני דוגמה לחיוג מהיר
INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description, DateAdded) VALUES (1, 'משרד', '031234567', 'מספר טלפון של המשרד', #2023-01-15#);

INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description, DateAdded) VALUES (2, 'בית', '098765432', 'מספר טלפון ביתי', #2023-02-20#);

INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description, DateAdded) VALUES (3, 'נייד', '0541234567', 'מספר טלפון נייד', #2023-03-10#);

INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description, DateAdded) VALUES (4, 'חירום', '100', 'משטרה - חירום', #2023-04-01#);

INSERT INTO SpeedDial (DialIndex, ContactName, PhoneNumber, Description, DateAdded) VALUES (5, 'מידע', '103', 'מידע טלפוני', #2023-05-05#);

-- הוספת אנשי קשר לדוגמה
INSERT INTO Contacts (ContactName, PhoneNumber, Landline, Email, Notes, DateAdded, CallCount) VALUES ('דוד כהן', '0521111111', NULL, 'david@email.com', 'עמית לעבודה', #2023-01-01#, 0);

INSERT INTO Contacts (ContactName, PhoneNumber, Landline, Email, Notes, DateAdded, CallCount) VALUES ('שרה לוי', '0532222222', NULL, 'sara@email.com', 'לקוח חשוב', #2023-02-01#, 0);

INSERT INTO Contacts (ContactName, PhoneNumber, Landline, Email, Notes, DateAdded, CallCount) VALUES ('יוסי אברהם', '0543333333', NULL, 'yossi@email.com', 'ספק', #2023-03-01#, 0);

INSERT INTO Contacts (ContactName, PhoneNumber, Landline, Email, Notes, DateAdded, CallCount) VALUES ('רחל ישראלי', '0504444444', NULL, 'rachel@email.com', 'חברה', #2023-04-01#, 0);

INSERT INTO Contacts (ContactName, PhoneNumber, Landline, Email, Notes, DateAdded, CallCount) VALUES ('משה דוד', '0555555555', NULL, 'moshe@email.com', 'שכן', #2023-05-01#, 0);
