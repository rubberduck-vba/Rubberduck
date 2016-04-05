using System;
using System.ComponentModel;
using System.Diagnostics;
using EventHook;
using EventHook.Hooks;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class MouseHookWrapper : IAttachable
    {
        //private IntPtr _hookId;

        //private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        //{
        //    if (nCode >= 0 && ((WM)wParam) == WM.RBUTTONDOWN)
        //    {
        //        OnMessageReceived();
        //    }

        //    return User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
        //}

        //public bool IsAttached { get; private set; }
        //public event EventHandler<HookEventArgs> MessageReceived;

        //private static IntPtr SetHook(User32.HookProc callback)
        //{
        //    var hook = User32.SetWindowsHookEx(WindowsHook.MOUSE_LL, callback, Kernel32.GetModuleHandle("user32"), 0);
        //    if (hook == IntPtr.Zero)
        //    {
        //        throw new Win32Exception();
        //    }
        //    return hook;
        //}

        //public void Attach()
        //{
        //    if (IsAttached)
        //    {
        //        return;
        //    }

        //    _hookId = SetHook(HookCallback);

        //    IsAttached = true;
        //    Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        //}

        //public void Detach()
        //{
        //    if (!IsAttached)
        //    {
        //        return;
        //    }

        //    User32.UnhookWindowsHookEx(_hookId);

        //    IsAttached = false;
        //    Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        //}
        public bool IsAttached { get; private set; }
        public event EventHandler<HookEventArgs> MessageReceived;
        public void Attach()
        {
            if (IsAttached)
            {
                return;
            }

            MouseWatcher.OnMouseInput += MouseWatcher_OnMouseInput;
            MouseWatcher.Start();
            IsAttached = true;
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            MouseWatcher.OnMouseInput -= MouseWatcher_OnMouseInput;
            MouseWatcher.Stop();
        }

        void MouseWatcher_OnMouseInput(object sender, MouseEventArgs e)
        {
            if (e.Message == MouseMessages.WM_RBUTTONDOWN)
            {
                OnMessageReceived();
            }
        }

        private void OnMessageReceived()
        {
            var handler = MessageReceived;
            if (handler != null)
            {
                handler.Invoke(this, HookEventArgs.Empty);
            }
        }
    }
}