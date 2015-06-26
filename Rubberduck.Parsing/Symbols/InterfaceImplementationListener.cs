using Rubberduck.Parsing.Grammar;
using Rubberduck.VBA;

namespace Rubberduck.Parsing.Symbols
{
    public class InterfaceImplementationListener : VBABaseListener
    {
        private readonly IdentifierReferenceResolver _resolver;

        public InterfaceImplementationListener(IdentifierReferenceResolver resolver)
        {
            _resolver = resolver;
        }

        public override void EnterImplementsStmt(VBAParser.ImplementsStmtContext context)
        {
            _resolver.Resolve(context);
        }

        public override void ExitModuleDeclarations(VBAParser.ModuleDeclarationsContext context)
        {
            throw new WalkerCancelledException();
        }
    }
}
