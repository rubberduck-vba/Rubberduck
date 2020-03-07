using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class MemberAccessMayReturnNothingInspectionBase : IdentifierReferenceInspectionFromDeclarationsBase
    {
        protected MemberAccessMayReturnNothingInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public abstract IEnumerable<Declaration> MembersUnderTest(DeclarationFinder finder);
        public abstract string ResultTemplate { get; }

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return MembersUnderTest(finder);
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            // prefilter to reduce search space
            if (reference.IsIgnoringInspectionResultFor(AnnotationName))
            {
                return false;
            }

            var usageContext = UsageContext(reference);

            var setter = usageContext is VBAParser.LExprContext lexpr 
                         && lexpr.Parent is VBAParser.SetStmtContext
                ? lexpr.Parent
                : null;

            if (setter is null)
            {
                return usageContext is VBAParser.MemberAccessExprContext 
                        || !(usageContext is VBAParser.CallStmtContext)
                            && !ContextIsNothingTest(usageContext);
            }

            var assignedTo = AssignmentTarget(reference, finder, setter);
            return assignedTo != null 
                   && IsUsedBeforeCheckingForNothing(assignedTo);
        }

        private static IdentifierReference AssignmentTarget(IdentifierReference reference, DeclarationFinder finder, ITree setter)
        {
            var assignedTo = finder.IdentifierReferences(reference.QualifiedModuleName)
                .SingleOrDefault(assign =>
                    assign.IsAssignment
                    && (assign.Context.GetAncestor<VBAParser.SetStmtContext>()?.Equals(setter) ?? false));
            return assignedTo;
        }

        private static RuleContext UsageContext(IdentifierReference reference)
        {
            var access = reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
            var usageContext = access.Parent is VBAParser.IndexExprContext indexExpr
                ? indexExpr.Parent
                : access.Parent;
            return usageContext;
        }

        private static bool ContextIsNothingTest(IParseTree context)
        {
            return context is VBAParser.LExprContext 
                   && context.Parent is VBAParser.RelationalOpContext comparison 
                   && comparison.IS() != null
                   && comparison.GetDescendent<VBAParser.ObjectLiteralIdentifierContext>() != null;
        }

        private static bool IsUsedBeforeCheckingForNothing(IdentifierReference assignedTo)
        {
            var tree = new Walker().GenerateTree(assignedTo.Declaration.ParentScopeDeclaration.Context, assignedTo.Declaration);
            var firstUse = GetReferenceNodes(tree).FirstOrDefault();

            return !(firstUse is null)
                   && !ContextIsNothingTest(firstUse.Reference.Context.Parent);
        }

        private static IEnumerable<INode> GetReferenceNodes(INode node)
        {
            if (node is ReferenceNode && node.Reference != null)
            {
                yield return node;
            }

            foreach (var childNode in node.Children.SelectMany(GetReferenceNodes))
            {
                yield return childNode;
            }
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var semiQualifiedName = $"{reference.Declaration.ParentDeclaration.IdentifierName}.{reference.IdentifierName}";
            return string.Format(ResultTemplate, semiQualifiedName);
        }
    }
}
