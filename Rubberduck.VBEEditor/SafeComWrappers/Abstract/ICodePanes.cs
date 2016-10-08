using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICodePanes : ISafeComWrapper, IComCollection<ICodePane>, IEquatable<ICodePanes>
    {
        IVBE Parent { get; }
        IVBE VBE { get; }
        ICodePane Current { get; set; }
    }
}