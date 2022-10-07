using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.InternalApi.Extensions;
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
    /// Warns when a variable is referenced prior to being assigned.
    /// </summary>
    /// <why>
    /// An uninitialized variable is being read, but since it's never assigned, the only value ever read would be the data type's default initial value. 
    /// Reading a variable that was never written to in any code path (especially if Option Explicit isn't specified), is likely to be a bug.
    /// </why>
    /// <remarks>
    /// This inspection may produce false positives when the variable is an array, or if it's passed by reference (ByRef) to a procedure that assigns it.
    /// </remarks>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     Debug.Print i ' i was never assigned
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     i = 42
    ///     Debug.Print i
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
    internal sealed class UnassignedVariableUsageInspection : IdentifierReferenceInspectionFromDeclarationsBase
    {
        public UnassignedVariableUsageInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        //See https://github.com/rubberduck-vba/Rubberduck/issues/2010 for why these are being excluded.
        private static readonly List<string> IgnoredFunctions = new List<string>
        {
            "VBE7.DLL;VBA.Strings.Len",
            "VBE7.DLL;VBA.Strings.LenB",
            "VBA6.DLL;VBA.Strings.Len",
            "VBA6.DLL;VBA.Strings.LenB"
        };

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Variable)
                .Where(declaration => !declaration.IsArray
                                      && !declaration.IsSelfAssigned
                                      && finder.MatchName(declaration.AsTypeName)
                                          .All(d => d.DeclarationType != DeclarationType.UserDefinedType)
                                      && !declaration.References
                                          .Any(reference => reference.IsAssignment)
                                      && !declaration.References
                                          .Any(reference => IsAssignedByRefArgument(reference.ParentScoping, reference, finder)));
        }

        //We override this in order to look up the argument usage exclusion references only once.
        protected override IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            var excludedReferenceSelections = DeclarationsWithExcludedArgumentUsage(finder)
                .SelectMany(SingleVariableArgumentSelections)
                .ToHashSet();

            return base.ObjectionableReferences(finder)
                .Where(reference => !excludedReferenceSelections.Contains(reference.QualifiedSelection)
                                    && !IsRedimedVariantArrayReference(reference));
        }

        private IEnumerable<ModuleBodyElementDeclaration> DeclarationsWithExcludedArgumentUsage(DeclarationFinder finder)
        {
            var vbaProjects = finder.Projects
                .Where(project => project.IdentifierName == "VBA" && !project.IsUserDefined)
                .ToList();

            if (!vbaProjects.Any())
            {
                return new List<ModuleBodyElementDeclaration>();
            }

            var stringModules = vbaProjects
                .Select(project => finder.FindStdModule("Strings", project, true))
                .OfType<ModuleDeclaration>()
                .ToList();

            if (!stringModules.Any())
            {
                return new List<ModuleBodyElementDeclaration>();
            }

            return stringModules
                .SelectMany(module => module.Members)
                .Where(decl => IgnoredFunctions.Contains(decl.QualifiedName.ToString()))
                .OfType<ModuleBodyElementDeclaration>();
        }

        private static IEnumerable<QualifiedSelection> SingleVariableArgumentSelections(ModuleBodyElementDeclaration member)
        {
            return member.Parameters
                .SelectMany(parameter => parameter.ArgumentReferences)
                .Select(SingleVariableArgumentSelection)
                .Where(maybeSelection => maybeSelection.HasValue)
                .Select(selection => selection.Value);
        }

        private static QualifiedSelection? SingleVariableArgumentSelection(ArgumentReference argumentReference)
        {
            var argumentContext = argumentReference.Context as VBAParser.LExprContext;
            if (!(argumentContext?.lExpression() is VBAParser.SimpleNameExprContext name))
            {
                return null;
            }

            return new QualifiedSelection(argumentReference.QualifiedModuleName, name.GetSelection());
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference != null
                   && !IsArraySubscriptAssignment(reference) 
                   && !IsArrayReDim(reference);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var identifierName = reference.IdentifierName;
            return string.Format(
                InspectionResults.UnassignedVariableUsageInspection,
                identifierName);
        }

        private static bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference, DeclarationFinder finder)
        {
            var argExpression = ImmediateArgumentExpressionContext(reference);

            if (argExpression is null)
            {
                return false;
            }

            var argument = argExpression.GetAncestor<VBAParser.ArgumentContext>();
            var parameter = finder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argument, enclosingProcedure);

            // note: not recursive, by design.
            return parameter != null
                && (parameter.IsImplicitByRef || parameter.IsByRef)
                && parameter.References.Any(r => r.IsAssignment);
        }

        private static VBAParser.ArgumentExpressionContext ImmediateArgumentExpressionContext(IdentifierReference reference)
        {
            var context = reference.Context;
            //The context is either already a simpleNameExprContext or an IdentifierValueContext used in a sub-rule of some other lExpression alternative. 
            var lExpressionNameContext = context is VBAParser.SimpleNameExprContext simpleName
                ? simpleName
                : context.GetAncestor<VBAParser.LExpressionContext>();

            //To be an immediate argument and, thus, assignable by ref, the structure must be argumentExpression -> expression -> lExpression.
            return lExpressionNameContext?
                .Parent?
                .Parent as VBAParser.ArgumentExpressionContext;
        }

        private static bool IsArraySubscriptAssignment(IdentifierReference reference)
        {
            var nameExpression = reference.Context;
            if (!(nameExpression.Parent is VBAParser.IndexExprContext indexExpression))
            {
                return false;
            }

            var callingExpression = indexExpression.Parent;

            return callingExpression is VBAParser.SetStmtContext 
                   || callingExpression is VBAParser.LetStmtContext;
        }

        private static bool IsArrayReDim(IdentifierReference reference)
        {
            var nameExpression = reference.Context;
            if (!(nameExpression.Parent is VBAParser.IndexExprContext indexExpression))
            {
                return false;
            }

            var reDimVariableStmt = indexExpression.Parent?.Parent;

            return reDimVariableStmt is VBAParser.RedimVariableDeclarationContext;
        }

        // This function works under the assumption that there are no assignments to the referenced variable.
        private bool IsRedimedVariantArrayReference(IdentifierReference reference)
        {
            if (reference.Declaration.AsTypeName != "Variant")
            {
                return false;
            }

            if(!reference.Context.TryGetAncestor<VBAParser.ModuleBodyElementContext>(out var containingMember))
            {
                return false;
            }

            var referenceSelection = reference.Selection;
            var referencedDeclarationName = reference.Declaration.IdentifierName;
            var reDimLocator = new PriorReDimLocator(referencedDeclarationName, referenceSelection);

            return reDimLocator.Visit(containingMember);
        }

        /// <summary>
        /// A visitor that visits a member's body and returns <c>true</c> if any <c>ReDim</c> statement for the variable called <c>name</c> is present before the <c>selection</c>.
        /// </summary>
        private class PriorReDimLocator : VBAParserBaseVisitor<bool>
        {
            private readonly string _name;
            private readonly Selection _selection;

            public PriorReDimLocator(string name, Selection selection)
            {
                _name = name;
                _selection = selection;
            }

            protected override bool DefaultResult => false;

            protected override bool ShouldVisitNextChild(Antlr4.Runtime.Tree.IRuleNode node, bool currentResult)
            {
                return !currentResult;
            }

            //This is actually the default implementation, but for explicities sake stated here.
            protected override bool AggregateResult(bool aggregate, bool nextResult)
            {
                return nextResult;
            }

            public override bool VisitRedimVariableDeclaration([NotNull] VBAParser.RedimVariableDeclarationContext context)
            {
                var reDimedVariableName = RedimedVariableName(context);
                if (reDimedVariableName != _name)
                {
                    return false;
                }

                var reDimSelection = context.GetSelection();

                return reDimSelection <= _selection;
            }

            private string RedimedVariableName([NotNull] VBAParser.RedimVariableDeclarationContext context)
            {
                if (!(context.expression() is VBAParser.LExprContext reDimmedVariablelExpr))
                {
                    //This is not syntactically correct VBA.
                    return null;
                }

                switch (reDimmedVariablelExpr.lExpression())
                {
                    case VBAParser.IndexExprContext indexExpr:
                        return indexExpr.lExpression().GetText();
                    case VBAParser.WhitespaceIndexExprContext whiteSpaceIndexExpr:
                        return whiteSpaceIndexExpr.lExpression().GetText();
                    default:  //This should not be possible in syntactically correct VBA.
                        return null;
                }
            }
        }
    }
}
