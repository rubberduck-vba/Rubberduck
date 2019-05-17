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
    /// Warns about assignments that appear to be assigning an object reference without the 'Set' keyword.
    /// </summary>
    /// <why>
    /// Omitting the 'Set' keyword will Let-coerce the right-hand side (RHS) of the assignment expression. If the RHS is an object variable,
    /// then the assignment is implicitly assigning to that object's default member, which may raise run-time error 91 at run-time.
    /// </why>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Object
    ///     foo = New Collection
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Object
    ///     Set foo = New Collection
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {

            return InterestingReferences().Select(reference =>
                new IdentifierReferenceInspectionResult(this,
                    string.Format(InspectionResults.ObjectVariableNotSetInspection, reference.Declaration.IdentifierName),
                    State, reference));
        }

        private IEnumerable<IdentifierReference> InterestingReferences()
        {
            var result = new List<IdentifierReference>();
            foreach (var moduleReferences in State.DeclarationFinder.IdentifierReferences())
            {
                var module = State.DeclarationFinder.ModuleDeclaration(moduleReferences.Key);
                if (module == null || !module.IsUserDefined || module.IsIgnoringInspectionResultFor(AnnotationName))
                {
                    // module isn't user code (?), or this inspection is ignored at module-level
                    continue;
                }

                result.AddRange(moduleReferences.Value.Where(reference => !reference.IsSetAssignment
                    && VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(reference, State)));
            }

            return result.Where(reference => !reference.IsIgnoringInspectionResultFor(AnnotationName));
        }
    }
}
