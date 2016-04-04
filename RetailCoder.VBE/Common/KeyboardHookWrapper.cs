using System;
using System.Diagnostics;
using EventHook;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Common
{
    public class KeyboardHookWrapper : IAttachable
    {
        private readonly VBE _vbe;

        public KeyboardHookWrapper(VBE vbe)
        {
            _vbe = vbe;
        }

        private int _lastLine;
        private void KeyboardWatcher_OnKeyInput(object sender, KeyInputEventArgs e)
        {
            var pane = _vbe.ActiveCodePane;
            if (pane == null)
            {
                return;
            }

            if (e.KeyData.EventType == KeyEvent.down)
            {
                // todo: handle keydown for 2-step hotkeys
                return;
            }

            int startLine;
            int endLine;
            int startColumn;
            int endColumn;
            // note: not using extension method because we don't need a QualifiedSelection here
            pane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

            var handler = MessageReceived;
            if (handler != null && _lastLine != startLine)
            {
                handler.Invoke(this, HookEventArgs.Empty);
                _lastLine = startLine;
            }
        }

        public bool IsAttached { get; private set; }
        public event EventHandler<HookEventArgs> MessageReceived;

        public void Attach()
        {
            KeyboardWatcher.Start();
            KeyboardWatcher.OnKeyInput += KeyboardWatcher_OnKeyInput;
            IsAttached = true;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }

        public void Detach()
        {
            KeyboardWatcher.Stop();
            KeyboardWatcher.OnKeyInput -= KeyboardWatcher_OnKeyInput;
            IsAttached = false;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }
    }
}