using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
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

        public override void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            _projectDeclarations.Clear();
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
                            _state.ParseTrees.Find(s => s.Key == module).Value,
                            token);
                    }
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException ?? exception; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }
        }
    }
}
