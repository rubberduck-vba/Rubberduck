using System;
using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.UI.Command
{
    public interface IMenuItem
    {
        string Key { get; }
        Func<string> Caption { get; }
        Func<string> ToolTipText { get; }
        bool BeginGroup { get; }
        int DisplayOrder { get; }
    }

    public interface ICommandMenuItem : IMenuItem
    {
        bool EvaluateCanExecute(RubberduckParserState state);
        bool HiddenWhenDisabled { get; }
        bool IsVisible { get; }
        ButtonStyle ButtonStyle { get; }
        CommandBase Command { get; }
        Image Image { get; }
        Image Mask { get; }
    }
}
