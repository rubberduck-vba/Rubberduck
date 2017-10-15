using System;
using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class ShowParserErrorsCommandMenuItem : CommandMenuItemBase
    {
        public ShowParserErrorsCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Error;
        }

        private string _tooltip;
        public void SetToolTip(string tooltip)
        {
            _tooltip = tooltip;
        }
        public override Func<string> ToolTipText { get { return () => _tooltip; } }

        public override bool IsVisible => false;
        public override bool HiddenWhenDisabled => true;
        public override ButtonStyle ButtonStyle => ButtonStyle.Icon;

        public override string Key => string.Empty;
        public override Image Image => Resources.cross_circle;
        public override Image Mask => Resources.circle_mask;
        public override int DisplayOrder => (int)RubberduckCommandBarItemDisplayOrder.ShowErrors;
    }
}