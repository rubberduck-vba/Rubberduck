using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public class ParserStateManager : ParserStateManagerBase
    {
        private const int _maxDegreeOfModuleStateChangeParallelism = -1;


        public ParserStateManager(RubberduckParserState state)
        : base(state) { }


        public override void SetModuleStates(ICollection<QualifiedModuleName> modules, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true)
        {
            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfModuleStateChangeParallelism;

            Parallel.ForEach(modules, options, module => SetModuleState(module, parserState, token, false));

            if (evaluateOverallParserState && !token.IsCancellationRequested)
            {
                EvaluateOverallParserState(token);
            }
        }

    }
}
