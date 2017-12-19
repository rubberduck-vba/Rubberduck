using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface ICommandBarControl : ISafeComWrapper, IEquatable<ICommandBarControl>
    {
        string Caption { get; set; }
        string DescriptionText { get; set; }
        string TooltipText { get; set; }
        string OnAction { get; set; }
        string Parameter { get; set; }
        string Tag { get; set; }
        bool BeginsGroup { get; set; }
        bool IsBuiltIn { get; }
        bool IsEnabled { get; set; }
        bool IsVisible { get; set; }
        bool IsPriorityDropped { get; }
        int Id { get; }
        int Index { get; }
        int Priority { get; set; }
        int Height { get; set; }
        int Width { get; set; }
        int Top { get; }
        int Left { get; }
        ControlType Type { get; }
        ICommandBar Parent { get; }

        void Execute();
        void Delete();
    }
}