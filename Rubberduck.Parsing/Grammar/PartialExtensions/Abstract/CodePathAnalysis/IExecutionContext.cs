namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    public interface IExecutionContext
    {
        void Assign(IAssignmentNode node);
        void EnterBranch(IBranchNode node);
        void EnterLoop(ILoopNode node);
        void EnterJump(IJumpNode node);

    }
}