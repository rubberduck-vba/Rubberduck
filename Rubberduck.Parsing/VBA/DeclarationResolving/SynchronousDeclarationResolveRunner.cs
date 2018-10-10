using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public class SynchronousDeclarationResolveRunner : DeclarationResolveRunnerBase
    {
        public SynchronousDeclarationResolveRunner(
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

            token.ThrowIfCancellationRequested();

            try
            {
                foreach(var module in modules)
                {
                    ResolveDeclarations(module, _state.ParseTrees.Find(s => s.Key == module).Value, projects, token);
                }
            }
            catch(OperationCanceledException)
            {
                throw;
            }
            catch (Exception)
            {
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }
        }
    }
}
