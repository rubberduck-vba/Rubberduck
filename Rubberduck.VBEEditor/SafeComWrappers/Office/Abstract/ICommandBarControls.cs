using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    // Abstraction of the CommandBarControls coclass interface in the interop assemblies for Office.v8 and Office.v12
    public interface ICommandBarControls : ISafeComWrapper, IComCollection<ICommandBarControl>, IEquatable<ICommandBarControls>
    {
        ICommandBar Parent { get; }
        ICommandBarButton AddButton(int? before = null);
        ICommandBarPopup AddPopup(int? before = null);
    }
}