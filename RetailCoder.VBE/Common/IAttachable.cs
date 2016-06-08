using System;

namespace Rubberduck.Common
{
    public interface IAttachable
    {
        bool IsAttached { get; }
        event EventHandler<HookEventArgs> MessageReceived;
        void Attach();
        void Detach();
    }
}
