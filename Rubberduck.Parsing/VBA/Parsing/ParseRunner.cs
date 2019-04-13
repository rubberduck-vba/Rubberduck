using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.Common;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class ParseRunner : ParseRunnerBase
    {
        private const int _maxDegreeOfParserParallelism = -1;

        public ParseRunner(
            RubberduckParserState state, 
            IParserStateManager parserStateManager, 
            IModuleParser parser) 
        :base(state, 
            parserStateManager, 
            parser)
        { }

        protected override IReadOnlyCollection<(QualifiedModuleName module, ModuleParseResults results)> ModulePareResults(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return new List<(QualifiedModuleName module, ModuleParseResults results)>();
            }

            token.ThrowIfCancellationRequested();

            var parsingStageTimer = ParsingStageTimer.StartNew();

            var results = new ConcurrentBag<(QualifiedModuleName module, ModuleParseResults results)>();

            var options = new ParallelOptions
            {
                CancellationToken = token,
                MaxDegreeOfParallelism = _maxDegreeOfParserParallelism
            };

            try
            {
                Parallel.ForEach(modules,
                    options,
                    module => results.Add((module, ModuleParseResults(module, token)))
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException ?? exception; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                StateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                throw;
            }

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Parsed user modules in {0}ms.");

            return results;
        }
    }
}
