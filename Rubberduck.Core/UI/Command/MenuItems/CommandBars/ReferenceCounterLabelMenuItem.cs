using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class ReferenceCounterLabelMenuItem : CommandMenuItemBase
    {
        public ReferenceCounterLabelMenuItem(FindAllReferencesCommand command)
            : base(command)
        {
            _caption = string.Empty;
        }

        private int _count;

        public void SetCaption(int referenceCount)
        {
            _count = referenceCount;
            _caption = $"{referenceCount} {RubberduckUI.ContextReferences}";
        }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return _count > 0;
        }

        private string _caption;
        public override Func<string> Caption { get { return () => _caption; } }

        public override string Key => string.Empty;
        public override bool BeginGroup => true;
        public override int DisplayOrder => (int)RubberduckCommandBarItemDisplayOrder.ContextRefCount;
    }
}