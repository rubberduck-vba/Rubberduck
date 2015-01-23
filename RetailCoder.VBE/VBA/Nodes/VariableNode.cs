using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class VariableDeclarationNode : Node
    {
        private readonly VisualBasic6Parser.VariableStmtContext _context;

        public VariableDeclarationNode(VisualBasic6Parser.VariableStmtContext context, string scope)
            :base(context, scope)
        {
            _context = context;

            foreach (var variable in context.variableListStmt().variableSubStmt())
            {
                AddChild(new VariableNode(variable, scope));
            }
        }
    }

    public class VariableNode : Node
    {
        public VariableNode(VisualBasic6Parser.VariableSubStmtContext context, string scope)
            : base(context, scope)
        {
        }
}
}