using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousParserStateManager:ParserStateManagerBase 
    {

        public SynchronousParserStateManager(RubberduckParserState state)
        : base(state) { } 

        public override void SetModuleStates(ICollection<QualifiedModuleName> modules, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true)
        {
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
