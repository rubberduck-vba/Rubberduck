using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.VBA
{
    public class ParseRunner : ParseRunnerBase
    {
        private const int _maxDegreeOfParserParallelism = -1;

        public ParseRunner(
            RubberduckParserState state, 
            IParserStateManager parserStateManager, 
            IStringParser parser,
            ISourceCodeProvider codePaneSourceCodeProvider,
            ISourceCodeProvider attributesSourceCodeProvider,
            IModuleRewriterFactory moduleRewriterFactory) 
        :base(state, 
            parserStateManager, 
            parser,
            codePaneSourceCodeProvider,
            attributesSourceCodeProvider,
            moduleRewriterFactory)
        { }

        public override void ParseModules(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            token.ThrowIfCancellationRequested();

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
        }
    }
}
