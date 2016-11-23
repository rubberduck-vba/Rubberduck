using System;
using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class ShowParserErrorsCommandMenuItem : CommandMenuItemBase
    {
        public ShowParserErrorsCommandMenuItem(CommandBase command) : base(command)
        {
        }

        private string _caption;
        public void SetCaption(string caption)
        {
            _caption = caption;
        }

        public override Func<string> Caption { get { return () => _caption; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Error;
        }

        //public override bool HiddenWhenDisabled { get { return true; } }
        public override bool BeginGroup { get { return true; } }

        public override string Key { get { return string.Empty; } }
        //public override Image Image { get { return Resources.cross_circle; } }
        //public override Image Mask { get { return Resources.cross_circle_mask; } }
        public override int DisplayOrder { get { return (int)RubberduckCommandBarItemDisplayOrder.ShowErrors; } }
    }
}