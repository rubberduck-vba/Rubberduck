using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousParserStateManager:ParserStateManagerBase 
    {
        public SynchronousParserStateManager(RubberduckParserState state)
        :base(state) { } 


        public override void SetModuleStates(IReadOnlyCollection<QualifiedModuleName> modules, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true)
        {
            if (!modules.Any())
            {
                return;
            }

            foreach (var module in modules)
            {
                SetModuleState(module, parserState, token, false);
            }

            if (evaluateOverallParserState && !token.IsCancellationRequested)
            {
                EvaluateOverallParserState(token);
            }
        }
    }
}
