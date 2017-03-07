using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class ContextSelectionLabelMenuItem : CommandMenuItemBase
    {
        public ContextSelectionLabelMenuItem()
            : base(null)
        {
            _caption = string.Empty;
        }

        private string _caption;
        public void SetCaption(string caption)
        {
            _caption = caption;
        }

        private string _tooltip;
        public void SetToolTip(string tooltip)
        {
            _tooltip = tooltip;
        }
        public override Func<string> ToolTipText { get { return () => _tooltip; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return false;
        }

        public override Func<string> Caption { get { return () => _caption; } }
        public override string Key { get { return string.Empty; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckCommandBarItemDisplayOrder.ContextStatus; } }
    }
}