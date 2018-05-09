using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Common
{
    /// <summary>
    /// Contains information about a captured key press resulting in modified code for a VBComponent.
    /// </summary>
    public class KeyHookEventArgs : EventArgs
    {
        public KeyHookEventArgs(Keys key, IVBComponent component)
        {
            Key = key;
            Component = component;
        }

        public Keys Key { get; }
        public IVBComponent Component { get; }
    }
}
