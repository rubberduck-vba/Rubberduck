using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface IControls : ISafeComWrapper, IComCollection<IControl>, IEquatable<IControls>
    {
        
    }
}