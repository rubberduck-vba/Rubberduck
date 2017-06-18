using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousReferenceResolveRunner : ReferenceResolveRunnerBase 
    {
        public SynchronousReferenceResolveRunner(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IModuleToModuleReferenceManager moduletToModuleReferenceManager,
            IReferenceRemover referenceRemover)
        : base(state,
            parserStateManager,
            moduletToModuleReferenceManager,
            referenceRemover)
        { }


        protected override void ResolveReferences(ICollection<KeyValuePair<QualifiedModuleName, IParseTree>> toResolve, CancellationToken token)
        {
            try
            {
                foreach(var kvp in toResolve)
                {
                    ResolveReferences(_state.DeclarationFinder, kvp.Key, kvp.Value, token);
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

        protected override void AddModuleToModuleReferences(DeclarationFinder finder, CancellationToken token)
        {
            try
            {
                foreach(var module in finder.AllModules())
                {
                    AddModuleToModuleReferences(finder, module);
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
