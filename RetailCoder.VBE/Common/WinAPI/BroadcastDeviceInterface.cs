using System;

namespace Rubberduck.Common.WinAPI
{
    struct BroadcastDeviceInterface
    {
        public Int32 DbccSize;
        public BroadcastDeviceType BroadcastDeviceType;
        public Int32 DbccReserved;
        public Guid DbccClassguid;
        public char DbccName;
    }
}
