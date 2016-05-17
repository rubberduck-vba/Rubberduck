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
        private readonly Keys _key;
        private readonly VBComponent _component;

        public KeyHookEventArgs(Keys key, VBComponent component)
        {
            _key = key;
            _component = component;
        }

        public Keys Key { get { return _key; } }
        public VBComponent Component { get { return _component; } }
    }
}
