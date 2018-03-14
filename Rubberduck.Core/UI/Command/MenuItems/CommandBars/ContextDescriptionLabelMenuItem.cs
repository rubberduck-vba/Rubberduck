using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class ContextDescriptionLabelMenuItem : CommandMenuItemBase
    {
        public ContextDescriptionLabelMenuItem()
            : base(null)
        {
        }

        public void SetCaption(string description)
        {
            _caption = description;
        }

        private string _caption;
        public override Func<string> Caption { get { return () => _caption; } }
        public override Func<string> ToolTipText { get { return () => _caption; } }

        public override string Key => string.Empty;
        public override bool EvaluateCanExecute(RubberduckParserState state) => false;

        public override int DisplayOrder => (int)RubberduckCommandBarItemDisplayOrder.ContextDescription;
    }
}