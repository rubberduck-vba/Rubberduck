using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IProperties : ISafeComWrapper, IComCollection<IProperty>, IEquatable<IProperties>
    {
        IVBE VBE { get; }
        IApplication Application { get; }
        object Parent { get; }

    }
}