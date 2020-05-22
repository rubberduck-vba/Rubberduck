using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about assignments that appear to be assigning an object reference without the 'Set' keyword.
    /// </summary>
    /// <why>
    /// Omitting the 'Set' keyword will Let-coerce the right-hand side (RHS) of the assignment expression. If the RHS is an object variable,
    /// then the assignment is implicitly assigning to that object's default member, which may raise run-time error 91 at run-time.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Object
    ///     foo = New Collection
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Object
    ///     Set foo = New Collection
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ObjectVariableNotSetInspection : IdentifierReferenceInspectionBase
    {
        public ObjectVariableNotSetInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            var failedLetCoercionAssignmentsInModule = FailedLetResolutionAssignments(module, finder);
            var possiblyObjectLhsLetAssignmentsWithFailedLetResolutionOnRhs = PossiblyObjectLhsLetAssignmentsWithNonValueOnRhs(module, finder);
            return failedLetCoercionAssignmentsInModule
                .Concat(possiblyObjectLhsLetAssignmentsWithFailedLetResolutionOnRhs);
        }

        private static IEnumerable<IdentifierReference> FailedLetResolutionAssignments(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.FailedLetCoercions(module)
                .Where(reference => reference.IsAssignment);
        }

        private static IEnumerable<IdentifierReference> PossiblyObjectLhsLetAssignmentsWithNonValueOnRhs(QualifiedModuleName module, DeclarationFinder finder)
        {
            return PossiblyObjectLhsLetAssignments(module, finder)
                .Where(tpl => finder.FailedLetCoercions(module)
                        .Any(reference => reference.Selection.Equals(tpl.rhs.GetSelection()))
                        || Tokens.Nothing.Equals(tpl.rhs.GetText(), StringComparison.InvariantCultureIgnoreCase))
                .Select(tpl => tpl.assignment);
        }

        private static IEnumerable<(IdentifierReference assignment, ParserRuleContext rhs)> PossiblyObjectLhsLetAssignments(QualifiedModuleName module, DeclarationFinder finder)
        {
            return PossiblyObjectNonSetAssignments(module, finder)
                .Select(reference => (reference, RhsOfLetAssignment(reference)))
                .Where(tpl => tpl.Item2 != null);
        }

        private static ParserRuleContext RhsOfLetAssignment(IdentifierReference letAssignment)
        {
            var letStatement = letAssignment.Context.Parent as VBAParser.LetStmtContext;
            return letStatement?.expression();
        }

        private static IEnumerable<IdentifierReference> PossiblyObjectNonSetAssignments(QualifiedModuleName module, DeclarationFinder finder)
        {
            var assignments = finder.IdentifierReferences(module)
                .Where(reference => reference.IsAssignment
                                    && !reference.IsSetAssignment
                                    && (reference.IsNonIndexedDefaultMemberAccess 
                                        || Tokens.Variant.Equals(reference.Declaration.AsTypeName, StringComparison.InvariantCultureIgnoreCase)));
            var unboundAssignments = finder.UnboundDefaultMemberAccesses(module)
                .Where(reference => reference.IsAssignment);

            return assignments.Concat(unboundAssignments);
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return true;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return string.Format(InspectionResults.ObjectVariableNotSetInspection, reference.IdentifierName);
        }
    }
}
