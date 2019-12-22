using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags publicly exposed instance fields.
    /// </summary>
    /// <why>
    /// Instance fields are the implementation details of a object's internal state; exposing them directly breaks encapsulation. 
    /// Often, an object only needs to expose a 'Get' procedure to expose an internal instance field.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Foo As Long
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Private internalFoo As Long
    /// 
    /// Public Property Get Foo() As Long
    ///     Foo = internalFoo
    /// End Property
    /// ]]>
    /// </example>
    public sealed class EncapsulatePublicFieldInspection : InspectionBase
    {
        public EncapsulatePublicFieldInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // we're creating a public field for every control on a form, needs to be ignored.
            var fields = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => item.DeclarationType != DeclarationType.Control
                               && (item.Accessibility == Accessibility.Public ||
                                   item.Accessibility == Accessibility.Global))
                .ToList();

            return fields
                .Select(issue => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionResults.EncapsulatePublicFieldInspection, issue.IdentifierName),
                                                      issue))
                .ToList();
        }
    }
}
