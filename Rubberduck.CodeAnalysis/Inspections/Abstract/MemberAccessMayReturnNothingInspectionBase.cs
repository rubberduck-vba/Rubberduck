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

        /// <summary>
        /// Members that might return Nothing
        /// </summary>
        /// <remarks>
        /// It must not be legal to call the members unqualified. In particular, user-defined members will not be considered.
        /// Moreover, this disqualifies all members on global objects. 
        /// </remarks>
        public abstract IEnumerable<Declaration> MembersUnderTest(DeclarationFinder finder);
        public abstract string ResultTemplate { get; }

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            //This restriction is in place because the inspection currently cannot handle unqualified accesses.
            return MembersUnderTest(finder).Where(member => !member.IsUserDefined);
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
                            && !ContextIsNothing(usageContext);
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
            //We prefer the with member access over the member access, because the accesses are resolved right to left.
            var access = reference.Context.GetAncestor<VBAParser.WithMemberAccessExprContext>() as VBAParser.LExpressionContext
                ?? reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();

            if (access == null)
            {
                return null;
            }

            return access.Parent is VBAParser.IndexExprContext indexExpr
                ? indexExpr.Parent
                : access.Parent;
        }

        private static bool ContextIsNothing(IParseTree context)
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
                   && !ContextIsNothing(firstUse.Reference.Context.Parent);
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
