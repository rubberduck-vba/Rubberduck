using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    // Abstraction of the CommandBar coclass interface in the interop assemblies for Office.v8 and Office.v12
    public interface ICommandBar : ISafeComWrapper, IEquatable<ICommandBar>
    {
        string Name { get; set; }
        int Id { get; }
        int Index { get; }
        int Top { get; set; }
        int Left { get; set; }
        int Width { get; set; }
        int Height { get; set; }
        bool IsBuiltIn { get; }
        bool IsEnabled { get; set; }
        bool IsVisible { get; set; }

        CommandBarType Type { get; }
        CommandBarPosition Position { get; set; }

        ICommandBarControls Controls { get; }

        ICommandBarControl FindControl(int id);
        ICommandBarControl FindControl(ControlType type, int id);

        void Delete();
    }
}