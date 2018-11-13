using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Abstract
{
   public abstract class MemberAccessMayReturnNothingInspectionBase : InspectionBase
    {
        protected MemberAccessMayReturnNothingInspectionBase(RubberduckParserState state) : base(state) { }

        public abstract List<Declaration> MembersUnderTest { get; }
        public abstract string ResultTemplate { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interesting = MembersUnderTest.SelectMany(member => member.References).ToList();
            if (!interesting.Any())
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var output = new List<IInspectionResult>();
            foreach (var reference in interesting.Where(use => !IsIgnoringInspectionResultFor(use, AnnotationName)))
            {
                var access = reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
                var usageContext = access.Parent is VBAParser.IndexExprContext
                    ? access.Parent.Parent
                    : access.Parent;

                var setter = usageContext is VBAParser.LExprContext lexpr && lexpr.Parent is VBAParser.SetStmtContext
                    ? lexpr.Parent
                    : null;

                if (setter is null)
                {
                    if (usageContext is VBAParser.MemberAccessExprContext || !ContextIsNothingTest(usageContext))
                    {
                        output.Add(new IdentifierReferenceInspectionResult(this,
                            string.Format(ResultTemplate,
                                $"{reference.Declaration.ParentDeclaration.IdentifierName}.{reference.IdentifierName}"),
                            State, reference));
                    }
                    continue;                   
                }

                var assignedTo = Declarations.SelectMany(decl => decl.References).SingleOrDefault(assign =>
                    assign.IsAssignment && (assign.Context.GetAncestor<VBAParser.SetStmtContext>()?.Equals(setter) ?? false));
                if (assignedTo is null)
                {
                    continue;                    
                }

                var tree = new Walker().GenerateTree(assignedTo.Declaration.ParentScopeDeclaration.Context, assignedTo.Declaration);
                var firstUse = GetReferenceNodes(tree).FirstOrDefault();
                if (firstUse is null || ContextIsNothingTest(firstUse.Reference.Context.Parent))
                {
                    continue;
                }

                output.Add(new IdentifierReferenceInspectionResult(this,
                    string.Format(ResultTemplate,
                        $"{reference.Declaration.ParentDeclaration.IdentifierName}.{reference.IdentifierName}"),
                    State, reference));
            }

            return output;
        }

        private bool ContextIsNothingTest(IParseTree context)
        {
            return context is VBAParser.LExprContext &&
                   context.Parent is VBAParser.RelationalOpContext comparison &&
                   comparison.IS() != null
                   && comparison.GetDescendent<VBAParser.ObjectLiteralIdentifierContext>() != null;
        }

        private IEnumerable<INode> GetReferenceNodes(INode node)
        {
            if (node is ReferenceNode && node.Reference != null)
            {
                yield return node;
            }
            
            foreach (var child in node.Children)
            {
                foreach (var childNode in GetReferenceNodes(child))
                {
                    yield return childNode;
                }
            }
        }
    }
}
