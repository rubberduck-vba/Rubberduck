using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class SynchronousParseRunner : ParseRunnerBase 
    {
        public SynchronousParseRunner(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IModuleParser parser) 
        :base(state, 
            parserStateManager,
            parser)
        { }


        protected override IReadOnlyCollection<(QualifiedModuleName module, ModuleParseResults results)> ModuleParseResults(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return new List<(QualifiedModuleName module, ModuleParseResults results)>();
            }

            token.ThrowIfCancellationRequested();

            var results = new List<(QualifiedModuleName module, ModuleParseResults results)>();

            try
            {
                foreach (var module in modules)
                {
                    results.Add((module, ModuleParseResults(module, token)));
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception)
            {
                StateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                throw;
            }

            return results;
        }
    }
}
