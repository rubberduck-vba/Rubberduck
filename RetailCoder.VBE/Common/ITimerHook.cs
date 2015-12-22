using System;

namespace Rubberduck.Common
{
    public interface ITimerHook : IAttachable
    {
        event EventHandler Tick;
    }
}