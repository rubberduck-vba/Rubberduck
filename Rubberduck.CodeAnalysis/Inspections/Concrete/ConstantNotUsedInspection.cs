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
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates 'Const' declarations that are never referenced.
    /// </summary>
    /// <why>
    /// Declarations that are never used should be removed.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Private Const foo As Long = 42
    ///
    /// Public Sub DoSomething()
    ///     ' no reference to 'foo' anywhere...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Private Const foo As Long = 42
    ///
    /// Public Sub DoSomething()
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ConstantNotUsedInspection : InspectionBase
    {
        public ConstantNotUsedInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = State.DeclarationFinder.UserDeclarations(DeclarationType.Constant)
                .Where(declaration => declaration.Context != null
                    && !declaration.References.Any())
                .ToList();

            return results.Select(issue => 
                new DeclarationInspectionResult(this,
                                     string.Format(InspectionResults.IdentifierNotUsedInspection, issue.DeclarationType.ToLocalizedString(), issue.IdentifierName),
                                     issue,
                                     new QualifiedContext<ParserRuleContext>(issue.QualifiedName.QualifiedModuleName, ((dynamic)issue.Context).identifier())));
        }
    }
}
