using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IControls : ISafeComWrapper, IComCollection<IControl>, IEquatable<IControls>
    {
        
    }
}