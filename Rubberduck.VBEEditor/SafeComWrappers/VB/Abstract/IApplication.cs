using System;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.Abstract
{
    public interface IApplication : ISafeComWrapper, IEquatable<IApplication>
    {
        string Version { get; }
    }
}