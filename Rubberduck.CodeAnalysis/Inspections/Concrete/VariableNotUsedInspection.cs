using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about variables that are never referenced.
    /// </summary>
    /// <why>
    /// A variable can be declared and even assigned, but if its value is never referenced, it's effectively an unused variable.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long ' declared
    ///     value = 42 ' assigned
    ///     ' ... but never rerenced
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class VariableNotUsedInspection : DeclarationInspectionBase
    {
        /// <summary>
        /// Inspection results for variables that are never referenced.
        /// </summary>
        /// <returns></returns>
        public VariableNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Variable)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            // exclude undeclared, see #5439
            return !declaration.IsWithEvents 
                   && !declaration.IsUndeclared
                   && declaration.References.All(reference => reference.IsAssignment)
                   && !declaration.References.Any(IsForLoopAssignment)
                   && !IsPublicInExposedClass(declaration);
        }

        private static bool IsPublicInExposedClass(Declaration procedure)
        {
            if (!(procedure.Accessibility == Accessibility.Public
                    || procedure.Accessibility == Accessibility.Global))
            {
                return false;
            }

            if (!(Declaration.GetModuleParent(procedure) is ClassModuleDeclaration classParent))
            {
                return false;
            }

            return classParent.IsExposed;
        }

        private bool IsForLoopAssignment(IdentifierReference reference)
        {
            if(!reference.IsAssignment)
            {
                return false;
            }

            //A For Next loop has the form For expr1 = expr2 (Step expr3) ... Next (expr1)
            var relationalOpAncestor = reference.Context.GetAncestor<VBAParser.RelationalOpContext>();
            return relationalOpAncestor != null
                && relationalOpAncestor.Parent is VBAParser.ForNextStmtContext;
        }

        protected override IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(declaration),
                declaration,
                Context(declaration));
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.IdentifierNotUsedInspection, 
                declarationType, 
                declarationName);
        }

        private QualifiedContext Context(Declaration declaration)
        {
            var module = declaration.QualifiedModuleName;
            var context = declaration.Context.GetDescendent<VBAParser.IdentifierContext>();
            return new QualifiedContext<ParserRuleContext>(module, context);
        }
    }
}
