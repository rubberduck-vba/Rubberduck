using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class ParseTreeInspectionBase : InspectionBase, IParseTreeInspection
    {
        protected ParseTreeInspectionBase(RubberduckParserState state)
            : base(state) { }

        public abstract IInspectionListener Listener { get; }
        public virtual ParsePass Pass => ParsePass.CodePanePass;
    }
}