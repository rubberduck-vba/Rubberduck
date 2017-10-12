using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Abstract
{
    public interface IControls : ISafeComWrapper, IComCollection<IControl>, IEquatable<IControls>
    {
        
    }
}