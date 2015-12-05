using System;

namespace Rubberduck.Common
{
    internal struct HookInfo
    {
        private readonly IntPtr _hookId;
        private readonly uint _keyCode;
        private readonly uint _shift;
        private readonly Action _action;

        public HookInfo(IntPtr hookId, uint keyCode, uint shift, Action action)
        {
            _hookId = hookId;
            _keyCode = keyCode;
            _shift = shift;
            _action = action;
        }

        public IntPtr HookId { get { return _hookId; } }
        public uint KeyCode { get { return _keyCode; } }
        public uint Shift { get { return _shift; } }
        public Action Action { get { return _action; } }
    }
}