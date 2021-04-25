using System.Runtime.InteropServices;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting.ComCommands
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodCommand : AddTestMethodBase
    {
        public AddTestMethodCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            IRewritingManager rewritingManager,
            ITestCodeGenerator codeGenerator, 
            IVbeEvents vbeEvents)
            : base(vbe, state, rewritingManager, codeGenerator, vbeEvents)
        {
            MethodGenerator = codeGenerator.GetNewTestMethodCode;
        }
    }
}
