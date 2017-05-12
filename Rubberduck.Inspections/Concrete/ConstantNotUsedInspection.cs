using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ConstantNotUsedInspection : InspectionBase
    {
        public ConstantNotUsedInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var results = State.DeclarationFinder.UserDeclarations(DeclarationType.Constant)
                .Where(declaration => declaration.Context != null
                    && !declaration.References.Any()
                    && !IsIgnoringInspectionResultFor(declaration, AnnotationName))
                .ToList();

            return results.Select(issue => 
                new DeclarationInspectionResult(this,
                                     string.Format(InspectionsUI.IdentifierNotUsedInspectionResultFormat, issue.DeclarationType.ToLocalizedString(), issue.IdentifierName),
                                     issue,
                                     new QualifiedContext<ParserRuleContext>(issue.QualifiedName.QualifiedModuleName, ((dynamic)issue.Context).identifier())));
        }
    }
}
