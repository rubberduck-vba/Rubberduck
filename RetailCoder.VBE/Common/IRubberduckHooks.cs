using System;
using System.Collections.Generic;

namespace Rubberduck.Common
{
    public interface IRubberduckHooks : IHook, IDisposable
    {
        IEnumerable<IAttachable> Hooks { get; }
        void AddHook(IAttachable hook);
    }
}