using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IAddIns : ISafeComWrapper, IComCollection<IAddIn>, IEquatable<IAddIns>
    {
        object Parent { get; }
        IVBE VBE { get; }
        void Update();
    }
}