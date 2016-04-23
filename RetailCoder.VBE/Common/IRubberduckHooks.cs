using System;
using System.Collections.Generic;

namespace Rubberduck.Common
{
    public interface IRubberduckHooks : IDisposable, IAttachable
    {
        void HookHotkeys();
    }
}