using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Results
{
    public class ProcedureCanBeWrittenAsFunctionInspectionResult : InspectionResultBase
    {
        private readonly Declaration _target;

        public ProcedureCanBeWrittenAsFunctionInspectionResult(IInspection inspection, RubberduckParserState state,
            Declaration target,
            QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext)
            : base(inspection, subStmtQualifiedContext.ModuleName, subStmtQualifiedContext.Context.subroutineName(), target)
        {
            _target =
                state.AllUserDeclarations.Single(declaration => declaration.DeclarationType == DeclarationType.Procedure &&
                                                                declaration.Context == subStmtQualifiedContext.Context);
        }
        
        public override string Description
        {
            get
            {
                return
                    string.Format(InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionResultFormat,
                        _target.IdentifierName).Capitalize();
            }
        }
    }
}
