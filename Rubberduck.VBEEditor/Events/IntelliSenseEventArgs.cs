using System;

namespace Rubberduck.VBEditor.Events
{
    public class IntelliSenseEventArgs : EventArgs
    {
        public static IntelliSenseEventArgs Shown => new IntelliSenseEventArgs(true);
        public static IntelliSenseEventArgs Hidden => new IntelliSenseEventArgs(false);
        internal IntelliSenseEventArgs(bool visible)
        {
            Visible = visible;
        }

        public bool Visible { get; }
    }
}
