using System;
using Rubberduck.VBEditor.SafeComWrappers.Forms;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface ICommandBarControls : ISafeComWrapper, IComCollection<ICommandBarControl>, IEquatable<ICommandBarControls>
    {
        ICommandBar Parent { get; }
        ICommandBarControl Add(ControlType type);
        ICommandBarControl Add(ControlType type, int before);
    }
}