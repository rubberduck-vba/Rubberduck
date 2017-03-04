using System;

namespace Rubberduck.Common
{
    public interface IRubberduckHooks : IDisposable, IAttachable
    {
        void HookHotkeys();
    }
}
