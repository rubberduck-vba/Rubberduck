using System;

namespace Rubberduck.Common
{
    public interface IHook : IAttachable
    {
        event EventHandler<HookEventArgs> MessageReceived;
    }
}