using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.CodePathAnalysis
{
    public class ProcedureTreeVisitorState
    {
        public ProcedureTreeVisitorState(IDeclarationFinderProvider finderProvider, QualifiedMemberName scope)
        {
            var finder = finderProvider.DeclarationFinder;

            Declarations = finder.AllDeclarations
                .Where(d => d.QualifiedName.Equals(scope) || (d.ParentScopeDeclaration?.QualifiedName.Equals(scope) ?? false))
                .ToImmutableHashSet();

            IdentifierReferences = finder.IdentifierReferences(scope)
                .ToImmutableHashSet();
        }

        public IEnumerable<Declaration> Declarations { get; }
        public IEnumerable<IdentifierReference> IdentifierReferences { get; }

        public INode CurrentNode { get; set; }

        private readonly Stack<GoSubJumpNode> _gosubNodes = new Stack<GoSubJumpNode>();

        private readonly IDictionary<Declaration, Stack<AssignmentNode>> _assignments
            = new Dictionary<Declaration, Stack<AssignmentNode>>();

        public IEnumerable<AssignmentNode> Assignments(Declaration declaration) => 
            _assignments.TryGetValue(declaration, out var value) 
                ? value 
                : Enumerable.Empty<AssignmentNode>();

        public void OnBranchNode(IParseTree tree)
        {
            CurrentNode = new BranchNode(tree);
        }

        public void OnLoopNode(IParseTree tree)
        {
            CurrentNode = new LoopNode(tree);
        }

        public void OnStatementNode(IParseTree tree)
        {
            CurrentNode = new StatementNode(tree);
        }

        public void OnBlockNode(IParseTree tree)
        {
            CurrentNode = new BlockNode(tree);
        }

        public void OnGoToNode(VBAParser.GoToStmtContext tree)
        {
            var targetName = tree.expression().GetText();
            var targetDeclaration = Declarations
                .SingleOrDefault(d =>
                    d.DeclarationType.HasFlag(DeclarationType.LineLabel)
                    && d.IdentifierName == targetName);
            var target = new LabelNode(targetDeclaration);
            var node = new GoToJumpNode(tree, target);

            CurrentNode = node;
        }

        public void OnGoSubNode(VBAParser.GoSubStmtContext tree)
        {
            var targetName = tree.expression().GetText();
            var targetDeclaration = Declarations
                .SingleOrDefault(d =>
                    d.DeclarationType.HasFlag(DeclarationType.LineLabel)
                    && d.IdentifierName == targetName);
            var target = new LabelNode(targetDeclaration);
            var node = new GoSubJumpNode(tree, target);
            _gosubNodes.Push(node);

            CurrentNode = node;
        }

        public void OnReturnNode(VBAParser.ReturnStmtContext tree)
        {
            var target = _gosubNodes.Count > 0
                ? _gosubNodes.Pop()
                : default;

            CurrentNode = new ReturnJumpNode(tree, target);
        }

        public OnErrorJumpNode CurrentErrorHandler { get; private set; }

        public void OnErrorNode(VBAParser.OnErrorStmtContext tree)
        {
            var name = tree.expression()?.GetText();
            var target = Declarations.SingleOrDefault(d => d.IdentifierName == name && d.DeclarationType.HasFlag(DeclarationType.LineLabel));
            var node = new OnErrorJumpNode(tree, new LabelNode(target));
            CurrentNode = node;
            CurrentErrorHandler = node;
        }

        public void OnResume(VBAParser.ResumeStmtContext tree)
        {
            var expression = tree.expression();
            if (expression != null)
            {
                var name = expression.GetText();
                var label = Declarations.SingleOrDefault(d => d.IdentifierName == name && d.DeclarationType.HasFlag(DeclarationType.LineLabel));
                var node = new ResumeJumpNode(tree, label?.Context);
                CurrentNode = node;
            }
            // todo: handle Resume Next?
        }

        /// <summary>
        /// Registers an assignment operation against a declaration.
        /// </summary>
        /// <param name="node"></param>
        public void OnAssignmentNode(AssignmentNode node)
        {
            var declaration = node.Reference.Declaration;
            if (!_assignments.TryGetValue(declaration, out var stack))
            {
                stack = new Stack<AssignmentNode>();
                _assignments[declaration] = stack;
            }
            stack.Push(node);
            CurrentNode = node;
        }

        /// <summary>
        /// Adds a usage to the last assignment of the declaration of the specified reference node.
        /// </summary>
        public void OnReferenceNode(ReferenceNode node)
        {
            var declaration = node.Reference.Declaration;
            if (!_assignments.TryGetValue(declaration, out var stack))
            {
                stack = new Stack<AssignmentNode>();
            }

            var assignment = stack.Count > 0
                ? stack.Peek()
                : null;

            assignment?.AddUsage(node);
            node.ValueNode = assignment;
            CurrentNode = node;
        }
    }
}