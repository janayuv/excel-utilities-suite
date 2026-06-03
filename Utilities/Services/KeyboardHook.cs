using System;
using System.Runtime.InteropServices;

namespace utilities.Services
{
    /// <summary>
    /// Intercepts Ctrl+Z at the OS level. When the add-in undo stack is non-empty the
    /// key is consumed and <see cref="Install"/>'s callback fires; otherwise the key
    /// falls through to Excel's native undo unchanged.
    /// </summary>
    internal static class KeyboardHook
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN    = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_Z          = 0x5A;
        private const int VK_CONTROL    = 0x11;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint   vkCode;
            public uint   scanCode;
            public uint   flags;
            public uint   time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        private static IntPtr               _hookId = IntPtr.Zero;
        private static LowLevelKeyboardProc _proc;   // keep delegate alive — prevents GC collection
        private static Action               _onCtrlZ;

        public static void Install(Action onCtrlZ)
        {
            if (_hookId != IntPtr.Zero) return;
            _onCtrlZ = onCtrlZ;
            _proc    = HookCallback;
            using (var proc   = System.Diagnostics.Process.GetCurrentProcess())
            using (var module = proc.MainModule)
            {
                _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _proc,
                    GetModuleHandle(module.ModuleName), 0);
            }

            if (_hookId == IntPtr.Zero)
                ErrorService.Log("WARN", "KeyboardHook: SetWindowsHookEx failed. Win32 error=" + Marshal.GetLastWin32Error());
            else
                ErrorService.Log("INFO", "KeyboardHook: installed (handle=" + _hookId + ")");
        }

        public static void Uninstall()
        {
            if (_hookId == IntPtr.Zero) return;
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 &&
                ((int)wParam == WM_KEYDOWN || (int)wParam == WM_SYSKEYDOWN))
            {
                var kb = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                bool isCtrlZ = kb.vkCode == VK_Z && (GetKeyState(VK_CONTROL) & 0x8000) != 0;

                if (isCtrlZ)
                {
                    ErrorService.Log("INFO", "KeyboardHook: Ctrl+Z detected. CanUndo=" + UndoService.CanUndo);
                    if (UndoService.CanUndo)
                    {
                        try { _onCtrlZ?.Invoke(); } catch (Exception ex) { ErrorService.Log("WARN", "KeyboardHook undo failed: " + ex.Message); }
                        return (IntPtr)1; // consumed — do not pass to Excel
                    }
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }
    }
}
