using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Highlights implicit Public access modifiers in user code.
    /// </summary>
    /// <why>
    /// In modern VB (VB.NET), the implicit access modifier is Private, as it is in most other programming languages.
    /// Making the Public modifiers explicit can help surface potentially unexpected language defaults.
    /// </why>
    /// <example>
    /// This inspection would flag the following implicit Public access modifier:
    /// <code>
    /// Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// </code>
    /// The following code would not trip the inspection:
    /// <code>
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// </code>
    /// </example>
    public sealed class ImplicitPublicMemberInspection : InspectionBase
    {
        public ImplicitPublicMemberInspection(RubberduckParserState state)
            : base(state) { }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionResults.ImplicitPublicMemberInspection, item.IdentifierName),
                                                     item);
            return issues;
        }
    }
}
