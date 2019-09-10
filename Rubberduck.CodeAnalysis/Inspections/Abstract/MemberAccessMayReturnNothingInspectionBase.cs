using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Inspections.Inspections.Extensions;
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
        private readonly ProcedureTreeVisitor _procedureTreeVisitor;

        protected MemberAccessMayReturnNothingInspectionBase(RubberduckParserState state,
            ProcedureTreeVisitor procedureTreeVisitor)
            : base(state)
        {
            _procedureTreeVisitor = procedureTreeVisitor;
        }

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

            var scopedRefGroups = interesting
                .Where(use => !use.IsIgnoringInspectionResultFor(AnnotationName))
                .GroupBy(r => r.ParentScoping.QualifiedName);

            foreach (var scopedRefGroup in scopedRefGroups)
            {
                var scope = scopedRefGroup.Key;
                var state = new ProcedureTreeVisitorState(State, scope);
                var tree = _procedureTreeVisitor.GenerateTree(scope, state);
                // todo: actually walk the tree
                foreach (var declarationGroup in scopedRefGroup.GroupBy(r => r.Declaration))
                {
                    var scopeName = declarationGroup.Key.ParentDeclaration.IdentifierName;
                    foreach (var reference in declarationGroup)
                    {
                        var access = reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
                        var usageContext = access.Parent is VBAParser.IndexExprContext
                            ? access.Parent.Parent
                            : access.Parent;

                        var setStmt = usageContext is VBAParser.LExprContext lexpr && lexpr.Parent is VBAParser.SetStmtContext
                            ? lexpr.Parent
                            : null;

                        if (setStmt is null)
                        {
                            if (usageContext is VBAParser.MemberAccessExprContext || !ContextIsNothingTest(usageContext))
                            {
                                output.Add(new IdentifierReferenceInspectionResult(this,
                                    string.Format(ResultTemplate, $"{scopeName}.{reference.IdentifierName}"), State, reference));
                            }

                            continue;
                        }

                        var assignedTo = Declarations.SelectMany(decl => decl.References).SingleOrDefault(assign =>
                            assign.IsAssignment && (assign.Context.GetAncestor<VBAParser.SetStmtContext>()?.Equals(setStmt) ?? false));
                        if (assignedTo is null)
                        {
                            continue;
                        }

                        var refs = GetReferenceNodes(tree, reference.Declaration);
                        var firstUse = refs.FirstOrDefault();
                        if (firstUse is null || ContextIsNothingTest(firstUse.Reference.Context.Parent))
                        {
                            continue;
                        }

                        output.Add(new IdentifierReferenceInspectionResult(this,
                            string.Format(ResultTemplate,
                                $"{reference.Declaration.ParentDeclaration.IdentifierName}.{reference.IdentifierName}"),
                            State, reference));
                    }
                }
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

        private IEnumerable<INode> GetReferenceNodes(INode node, Declaration variable)
        {
            if (node.ParseTree is ITerminalNode)
            {
                yield break;
            }

            if (node is ReferenceNode && node.Reference != null && node.Reference.Declaration.Equals(variable))
            {
                yield return node;
            }

            foreach (var child in node.Children)
            {
                foreach (var childNode in GetReferenceNodes(child, variable))
                {
                    yield return childNode;
                }
            }
        }
    }
}
