using System;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Common
{
    /// <summary>
    /// Contains information about a captured key press resulting in modified code for a VBComponent.
    /// </summary>
    public class KeyHookEventArgs : EventArgs
    {
        public KeyHookEventArgs(Keys key, VBComponent component)
        {
            Key = key;
            Component = component;
        }

        public Keys Key { get; }
        public VBComponent Component { get; }
    }
}
