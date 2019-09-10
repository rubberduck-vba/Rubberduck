using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class LabelNode : NodeBase
    {
        public LabelNode(Declaration declaration) : base(declaration.Context)
        {
            Declaration = declaration;
        }

        public Declaration Declaration { get; }
    }

    public abstract class JumpNode<T> : StatementNode 
        where T : IParseTree
    {
        public JumpNode(T tree, IParseTree target) : base(tree)
        {
            Target = target;
        }

        public IParseTree Target { get; }
    }

    public class GoToJumpNode : JumpNode<VBAParser.GoToStmtContext>
    {
        public GoToJumpNode(VBAParser.GoToStmtContext tree, LabelNode target)
            : base(tree, target.ParseTree)
        {

        }

        public LabelNode Target { get; }
    }

    public class GoSubJumpNode : JumpNode<VBAParser.GoSubStmtContext>
    {
        public GoSubJumpNode(VBAParser.GoSubStmtContext tree, LabelNode target) 
            : base(tree, target.ParseTree) { }
    }

    public class ReturnJumpNode : JumpNode<VBAParser.ReturnStmtContext>
    {
        public ReturnJumpNode(VBAParser.ReturnStmtContext tree, JumpNode<VBAParser.GoSubStmtContext> origin) 
            : base(tree, origin.ParseTree) { }

        public bool HasReturnTarget => Target != null; // if false, that's "return without gosub" run-time error 3.
    }

    public class OnErrorJumpNode : JumpNode<VBAParser.OnErrorStmtContext>
    {
        public OnErrorJumpNode(VBAParser.OnErrorStmtContext tree, LabelNode target) 
            : base(tree, target?.ParseTree) { }
    }

    public class ResumeJumpNode : JumpNode<VBAParser.ResumeStmtContext>
    {
        public ResumeJumpNode(VBAParser.ResumeStmtContext tree, IParseTree target)
            : base(tree, target) { }
    }
}