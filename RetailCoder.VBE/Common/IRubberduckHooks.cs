using System;
using System.Collections.Generic;

namespace Rubberduck.Common
{
    public interface IRubberduckHooks : IHook, IDisposable
    {
        IEnumerable<IHook> Hooks { get; }
        void AddHook<THook>(THook hook) where THook : IHook;
    }
}