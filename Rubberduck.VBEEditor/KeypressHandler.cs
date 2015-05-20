using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor
{
    // Thank you, http://stackoverflow.com/questions/8120199/how-to-get-the-keypress-event-from-a-word-2010-addin-developed-in-c/8800908#8800908
    public partial class ActiveCodePaneEditor : IActiveCodePaneEditor
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private static IntPtr hookId = IntPtr.Zero;
        private delegate IntPtr HookProcedure(int nCode, IntPtr wParam, IntPtr lParam);
        private static HookProcedure procedure = HookCallback;

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProcedure lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr SetHook(HookProcedure procedure)
        {
            using (Process process = Process.GetCurrentProcess())
            using (ProcessModule module = process.MainModule)
                return SetWindowsHookEx(WH_KEYBOARD_LL, procedure, GetModuleHandle(module.ModuleName), 0);
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int pointerCode = Marshal.ReadInt32(lParam);
                string pressedKey = ((Keys)pointerCode).ToString();

                //Do some sort of processing on key press
                var thread = new Thread(() => { MessageBox.Show(pressedKey); });
                thread.Start();
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        private void InternalStartup()
        {
            hookId = SetHook(procedure);
        }
    }
}
