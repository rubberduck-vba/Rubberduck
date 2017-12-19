using System;
using System.Windows.Forms;

namespace Rubberduck.Common
{
    public class HookEventArgs : EventArgs
    {
        private static readonly Lazy<HookEventArgs> _empty = new Lazy<HookEventArgs>(() => new HookEventArgs(Keys.None));

        public HookEventArgs(Keys key)
        {
            Key = key;
        }

        public Keys Key { get; }

        public new static HookEventArgs Empty => _empty.Value;
    }
}
