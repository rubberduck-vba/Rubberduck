using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Rubberduck.Parsing.VBA
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


        public override void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            _projectDeclarations.Clear();
            token.ThrowIfCancellationRequested();

            try
            {
                foreach(var module in modules)
                {
                    ResolveDeclarations(module, _state.ParseTrees.Find(s => s.Key == module).Value, token);
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
