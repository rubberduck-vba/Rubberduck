using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class ParseTreeInspectionBase : InspectionBase, IParseTreeInspection
    {
        protected ParseTreeInspectionBase(RubberduckParserState state)
            : base(state) { }

        public abstract IInspectionListener Listener { get; }
        public virtual CodeKind TargetKindOfCode => CodeKind.CodePaneCode;
    }
}