using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IReferences : ISafeComWrapper, IComCollection<IReference>, IEquatable<IReferences>
    {
        event EventHandler<ReferenceEventArgs> ItemAdded;
        event EventHandler<ReferenceEventArgs> ItemRemoved;

        IVBE VBE { get; }
        VBProject Parent { get; }

        IReference AddFromGuid(string guid, int major, int minor);
        IReference AddFromFile(string path);
        void Remove(IReference reference);
    }
}