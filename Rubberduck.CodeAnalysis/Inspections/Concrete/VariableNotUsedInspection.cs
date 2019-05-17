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
    /// Warns about variables that are never referenced.
    /// </summary>
    /// <why>
    /// A variable can be declared and even assigned, but if its value is never referenced, it's effectively an unused variable.
    /// </why>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long ' declared
    ///     value = 42 ' assigned
    ///     ' ... but never rerenced
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class VariableNotUsedInspection : InspectionBase
    {
        public VariableNotUsedInspection(RubberduckParserState state) : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(declaration =>
                    !declaration.IsWithEvents
                    && !declaration.IsIgnoringInspectionResultFor(AnnotationName)
                    && !declaration.References.Any());

            return declarations.Select(issue => 
                new DeclarationInspectionResult(this,
                                     string.Format(InspectionResults.IdentifierNotUsedInspection, issue.DeclarationType.ToLocalizedString(), issue.IdentifierName),
                                     issue,
                                     new QualifiedContext<ParserRuleContext>(issue.QualifiedName.QualifiedModuleName, ((dynamic)issue.Context).identifier())));
        }
    }
}
