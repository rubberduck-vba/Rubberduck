using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IApplication : ISafeComWrapper, IEquatable<IApplication>
    {
        string Version { get; }
    }
}