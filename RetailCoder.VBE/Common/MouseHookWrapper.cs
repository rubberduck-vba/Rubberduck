using System;
using EventHook;
using EventHook.Hooks;

namespace Rubberduck.Common
{
    public class MouseHookWrapper : IAttachable
    {
        public MouseHookWrapper()
        {
            MouseWatcher.OnMouseInput += MouseWatcher_OnMouseInput;
        }

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
            IsAttached = true;
        }

        public void Detach()
        {
            MouseWatcher.Stop();
            IsAttached = false;
        }
    }
}