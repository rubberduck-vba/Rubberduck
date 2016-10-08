using System;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract
{
    public interface IApplication : ISafeComWrapper, IEquatable<IApplication>
    {
        string Version { get; }
    }
}