using System;
using System.Diagnostics;
using EventHook;
using EventHook.Hooks;

namespace Rubberduck.Common
{
    public class MouseHookWrapper : IAttachable
    {
        private void MouseWatcher_OnMouseInput(object sender, MouseEventArgs e)
        {
            // only handle right-clicks
            if (e.Message != MouseMessages.WM_RBUTTONDOWN)
            {
                return;
            }

            var handler = MessageReceived;
            if (handler != null)
            {
                handler.Invoke(this, HookEventArgs.Empty);
            }
        }

        public bool IsAttached { get; private set; }
        public event EventHandler<HookEventArgs> MessageReceived;

        public void Attach()
        {
            MouseWatcher.Start();
            MouseWatcher.OnMouseInput += MouseWatcher_OnMouseInput;
            IsAttached = true;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }

        public void Detach()
        {
            MouseWatcher.Stop();
            MouseWatcher.OnMouseInput -= MouseWatcher_OnMouseInput;
            IsAttached = false;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }
    }
}