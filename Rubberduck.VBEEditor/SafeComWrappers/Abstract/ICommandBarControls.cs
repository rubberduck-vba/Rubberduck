using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICommandBarControls : ISafeComWrapper, IComCollection<ICommandBarControl>, IEquatable<ICommandBarControls>
    {
        ICommandBar Parent { get; }
        ICommandBarControl Add(ControlType type);
        ICommandBarControl Add(ControlType type, int before);
    }
}