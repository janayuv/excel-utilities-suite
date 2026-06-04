using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace utilities.Services
{
    /// <summary>
    /// Intercepts Ctrl+Z at the OS level. When the add-in undo stack is non-empty the
    /// key is consumed and <see cref="Install"/>'s callback fires; otherwise the key
    /// falls through to Excel's native undo unchanged.
    ///
    /// SAFETY RULE: the hook callback must return in &lt;200 ms or Windows kills it and
    /// can crash the host process. Therefore the callback does nothing but a pure
    /// in-memory check (<see cref="UndoService.CanUndo"/>), then posts the actual undo
    /// work to the UI thread via <see cref="SynchronizationContext"/> so all COM calls
    /// happen outside the hook dispatch window.
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

        private static IntPtr                _hookId = IntPtr.Zero;
        private static LowLevelKeyboardProc  _proc;         // keep delegate alive — prevents GC collection
        private static Action                _onCtrlZ;
        private static SynchronizationContext _syncContext; // captured on UI thread at Install time

        public static void Install(Action onCtrlZ)
        {
            if (_hookId != IntPtr.Zero) return;

            // Capture the UI-thread SynchronizationContext BEFORE installing the hook.
            // All COM calls will be posted back here, outside the hook callback window.
            _syncContext = SynchronizationContext.Current;
            _onCtrlZ     = onCtrlZ;
            _proc        = HookCallback;

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
            // CRITICAL: no I/O, no COM, no blocking calls here.
            // This callback runs inside Windows' input hook dispatch — any delay or
            // re-entrant COM call can deadlock or crash Excel within ~200 ms.
            if (nCode >= 0 &&
                ((int)wParam == WM_KEYDOWN || (int)wParam == WM_SYSKEYDOWN))
            {
                var kb = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));

                if (kb.vkCode == VK_Z && (GetKeyState(VK_CONTROL) & 0x8000) != 0
                    && UndoService.CanUndo)   // pure in-memory check — safe inside hook
                {
                    // Post the undo to the UI message loop so COM calls happen
                    // after this hook callback has already returned to Windows.
                    _syncContext?.Post(_ =>
                    {
                        try { _onCtrlZ?.Invoke(); }
                        catch (Exception ex) { ErrorService.Log("WARN", "KeyboardHook undo failed: " + ex.Message); }
                    }, null);

                    return (IntPtr)1; // consumed — do not pass to Excel
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }
    }
}
