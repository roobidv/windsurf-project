# Changelog - April 26, 2026

## Session Summary
Backup: `C:\Temp\BAS_Backup_20260426_2111\` (28 .bas files)

---

## New Features

### 1. Restore Tables from Backup (`RestoreBackup.bas`) — NEW MODULE
- **Commands**: `#שחזר#` or `#שיחזור#` in txtSearch
- Finds the **newest .accdb** file in `C:\Temp\`
- Confirmation dialog showing file name and date
- Restores 6 tables: `Contacts`, `CallHistory`, `SpeedDial`, `Interactions`, `tblSettings`, `tblGLOBAL_PHONE_BOOK`
- Skips AutoNumber fields, reports how many tables restored
- Added to `ModuleUpdater2.bas` for auto-update

### 2. SyncPhoneBook Moved to Form_Unload
- `RunSyncPhoneBook` now runs on **form close** instead of form load
- No impact on startup time — sync happens when user closes the app
- Ping check changed from WMI (slow ~3s) to `ping -n 1 -w 500` (fast ~0.5s)

---

## Bug Fixes

### 3. Space Key Not Working in txtSearch (CRITICAL FIX)
**Problem**: Typing space (especially as 3rd character) caused it to be deleted and cursor to jump.

**Root Cause**: `ContactsDialer_Form_KeyDown` processed ALL keys through multiple `GetAsyncKeyState` checks, and the `TxtSearch_Change` event triggered cascading events (`AfterUpdate` → `LoadSelectedContact` → `RefreshCallHistoryGrid`) that stole focus.

**Fix** (3 parts):
1. **Guard in KeyDown**: When txtSearch has focus, immediately `Exit Function` for non-special keys (space, letters, etc.). Only process: Enter, Down, ESC, F8, F9, ALT+n, Ctrl+key.
2. **Cached key states**: `GetAsyncKeyState` called once per key (`kDown`, `kEnter`, `kEsc`) and reused — prevents double-read issues on quick key taps.
3. **Change event restructured**: Explicitly calls `LoadSelectedContact` and `RefreshCallHistoryGrid`, then restores focus with `SetFocus` + `SelStart` AFTER grid refresh.

### 4. Removed Blanket `ContactsDialer_Form_KeyDown = True`
- Previously, the function returned `True` for ALL keys at the end (line 994)
- This could suppress unhandled keys — removed so only handled keys return True

---

## Files Modified
| File | Change |
|------|--------|
| `ContactsDialerCode.bas` | Guard in KeyDown, cached key states, Change event restructured, SyncPhoneBook moved to Unload |
| `SyncPhoneBook.bas` | Ping changed from WMI to shell `ping -w 500` |
| `RestoreBackup.bas` | **NEW** — restore tables from newest backup |
| `ModuleUpdater2.bas` | Added `DoUpdate "RestoreBackup"` |
| `mdlPhone.bas` | Removed sync timer logic (`m_syncDone`, `syncDelay`) |

---

## Commands Reference
| Command | Action |
|---------|--------|
| `#גיבוי#` / `#גבוי#` | Force backup to `C:\Temp\` |
| `#שחזר#` / `#שיחזור#` | Restore tables from newest backup in `C:\Temp\` |
