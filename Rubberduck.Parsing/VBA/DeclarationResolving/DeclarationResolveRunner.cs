using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public class DeclarationResolveRunner : DeclarationResolveRunnerBase
    {
        private const int _maxDegreeOfDeclarationResolverParallelism = -1;

        public DeclarationResolveRunner(
            RubberduckParserState state, 
            IParserStateManager parserStateManager, 
            IProjectReferencesProvider projectReferencesProvider) 
        :base(
            state, 
            parserStateManager, 
            projectReferencesProvider)
        { }

        protected override void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, IDictionary<string, ProjectDeclaration> projects, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            var parsingStageTimer = ParsingStageTimer.StartNew();

            token.ThrowIfCancellationRequested();

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfDeclarationResolverParallelism;
            try
            {
                Parallel.ForEach(modules,
                    options,
                    module =>
                    {
                        ResolveDeclarations(module,
                            _state.GetParseTree(module),
                            _state.GetLogicalLines(module),
                            projects,
                            token);
                    }
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    //This rethrows the exception with the original stack trace.
                    ExceptionDispatchInfo.Capture(exception.InnerException ?? exception).Throw();
                }
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Resolved user declaration in {0}ms.");
        }
    }
}
