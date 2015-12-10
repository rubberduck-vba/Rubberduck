using System;
using System.Windows.Forms;
using Rubberduck.Annotations;

namespace Rubberduck.Common
{
    public struct HookInfo
    {
        private readonly IntPtr _hookId;
        private readonly Keys _key;
        private readonly uint _shift;
        private readonly Action _action;

        public HookInfo(IntPtr hookId, Keys key, uint shift, Action action)
        {
            _hookId = hookId;
            _key = key;
            _shift = shift;
            _action = action;
        }

        public IntPtr HookId { get { return _hookId; } }
        public Keys Key { get { return _key; } }
        public uint Shift { get { return _shift; } }
        public Action Action { get { return _action; } }
    }
}