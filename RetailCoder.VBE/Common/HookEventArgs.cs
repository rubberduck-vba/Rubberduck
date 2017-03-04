using System;
using System.Windows.Forms;

namespace Rubberduck.Common
{
    public class HookEventArgs : EventArgs
    {
        private readonly Keys _key;
        private static readonly Lazy<HookEventArgs> _empty = new Lazy<HookEventArgs>(() => new HookEventArgs(Keys.None));

        public HookEventArgs(Keys key)
        {
            _key = key;
        }

        public Keys Key { get { return _key; } }

        public new static HookEventArgs Empty {get { return _empty.Value; }}
    }
}
