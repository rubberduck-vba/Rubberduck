using System;

namespace Rubberduck.AddRemoveReferences
{
    [Flags]
    public enum ReferenceStatus
    {
        None = 0,
        BuiltIn = 1 << 1,
        Loaded = 1 << 2,
        Broken = 1 << 3,
        Pinned = 1 << 4,
        Recent = 1 << 5,
        Added = 1 << 6
    }
}