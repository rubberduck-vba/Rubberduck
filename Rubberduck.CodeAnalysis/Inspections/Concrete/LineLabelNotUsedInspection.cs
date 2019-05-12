using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class LineLabelNotUsedInspection : InspectionBase
    {
        public LineLabelNotUsedInspection(RubberduckParserState state) : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var labels = State.DeclarationFinder.UserDeclarations(DeclarationType.LineLabel);
            var declarations = labels
                .Where(declaration =>
                    !declaration.IsWithEvents
                    && declaration.Context is IdentifierStatementLabelContext
                    && !declaration.IsIgnoringInspectionResultFor(AnnotationName)
                    && (!declaration.References.Any() || declaration.References.All(reference => reference.IsAssignment)));

            return declarations.Select(issue => 
                new DeclarationInspectionResult(this,
                                     string.Format(InspectionResults.IdentifierNotUsedInspection, issue.DeclarationType.ToLocalizedString(), issue.IdentifierName),
                                     issue,
                                     new QualifiedContext<ParserRuleContext>(issue.QualifiedName.QualifiedModuleName, ((IdentifierStatementLabelContext)issue.Context).legalLabelIdentifier())));
        }
    }
}
