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

        private string _caption;
        private int _count;

        public void SetCaption(int referenceCount)
        {
            _count = referenceCount;
            _caption = string.Format("{0} {1}", referenceCount, RubberduckUI.ContextReferences);
        }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return _count > 0;
        }

        public override Func<string> Caption { get { return () => _caption; } }
        public override string Key { get { return string.Empty; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckCommandBarItemDisplayOrder.ContextRefCount; } }
    }
}