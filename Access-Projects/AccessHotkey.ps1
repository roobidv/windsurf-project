Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;

public class AccessHotkey : Form {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool OpenIcon(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
    [DllImport("user32.dll")]
    public static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);
    [DllImport("user32.dll")]
    public static extern IntPtr ActivateKeyboardLayout(IntPtr hkl, uint Flags);
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    const int WM_HOTKEY = 0x0312;
    const int SW_RESTORE = 9;
    const uint VK_F2 = 0x71;

    private System.Windows.Forms.Timer watchdog;

    public AccessHotkey() {
        this.ShowInTaskbar = false;
        this.WindowState = FormWindowState.Minimized;
        this.Visible = false;
        this.FormBorderStyle = FormBorderStyle.None;
        this.Size = new System.Drawing.Size(1, 1);
        RegisterHotKey(this.Handle, 1, 0, VK_F2);

        // Auto-exit when Access closes
        watchdog = new System.Windows.Forms.Timer();
        watchdog.Interval = 5000;
        watchdog.Tick += (s, e) => {
            if (Process.GetProcessesByName("MSACCESS").Length == 0)
                Application.Exit();
        };
        watchdog.Start();
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            {
                IntPtr h = IntPtr.Zero;
                foreach (var p in Process.GetProcessesByName("MSACCESS")) {
                    h = p.MainWindowHandle;
                    break;
                }
                if (h == IntPtr.Zero) h = FindWindow("OMain", null);
                if (h != IntPtr.Zero) {
                    // Try to find the form window directly
                    IntPtr hForm = FindWindow(null, "Contacts Dialer");
                    if (hForm != IntPtr.Zero) {
                        ShowWindow(hForm, SW_RESTORE);
                        SetForegroundWindow(hForm);
                    } else {
                        ShowWindow(h, SW_RESTORE);
                        SetForegroundWindow(h);
                    }
                    Thread.Sleep(150);
                    // Switch to Hebrew keyboard
                    IntPtr hkl = LoadKeyboardLayout("0000040D", 1);
                    ActivateKeyboardLayout(hkl, 0);
                    // Signal VBA to focus txtSearch
                    string flagPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "hotkey_flag.txt");
                    System.IO.File.WriteAllText(flagPath, "1");
                }
            }
        }
        base.WndProc(ref m);
    }

    protected override void OnFormClosing(FormClosingEventArgs e) {
        watchdog.Stop();
        UnregisterHotKey(this.Handle, 1);
        base.OnFormClosing(e);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing

$form = New-Object AccessHotkey
[System.Windows.Forms.Application]::Run($form)
