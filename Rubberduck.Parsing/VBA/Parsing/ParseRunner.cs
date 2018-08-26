using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class ParseRunner : ParseRunnerBase
    {
        private const int _maxDegreeOfParserParallelism = -1;

        private static Logger Logger = LogManager.GetCurrentClassLogger();

        public ParseRunner(
            RubberduckParserState state, 
            IParserStateManager parserStateManager, 
            IModuleParser parser) 
        :base(state, 
            parserStateManager, 
            parser)
        { }

        public override void ParseModules(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            token.ThrowIfCancellationRequested();

            var stopwatch = Stopwatch.StartNew();

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfParserParallelism;

            try
            {
                Parallel.ForEach(modules,
                    options,
                    module => ParseModule(module, token)
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

            stopwatch.Stop();
            Logger.Debug($"Parsed user modules in {stopwatch.ElapsedMilliseconds}ms.");
        }
    }
}
