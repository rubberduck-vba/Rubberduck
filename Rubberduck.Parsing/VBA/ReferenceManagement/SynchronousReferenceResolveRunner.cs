using System;
using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public class SynchronousReferenceResolveRunner : ReferenceResolveRunnerBase 
    {
        public SynchronousReferenceResolveRunner(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager,
            IReferenceRemover referenceRemover,
            IDocumentModuleSuperTypeNamesProvider documentModuleSuperTypeNamesProvider)
        : base(state,
            parserStateManager,
            moduleToModuleReferenceManager,
            referenceRemover,
            documentModuleSuperTypeNamesProvider)
        {}


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
                foreach(var module in finder.AllModules)
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
