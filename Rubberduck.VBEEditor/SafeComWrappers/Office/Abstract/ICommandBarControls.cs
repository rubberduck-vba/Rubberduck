using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Abstract
{
    public interface ICommandBarControls : ISafeComWrapper, IComCollection<ICommandBarControl>, IEquatable<ICommandBarControls>
    {
        ICommandBar Parent { get; }
        ICommandBarControl Add(ControlType type);
        ICommandBarControl Add(ControlType type, int before);
    }
}