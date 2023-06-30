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
                    ResolveDeclarations(
                        module, 
                        _state.GetParseTree(module),
                        _state.GetLogicalLines(module),
                        projects,
                        token);
                    var declaration = _state.DeclarationFinder.ModuleDeclaration(module);
                    if (declaration is DocumentModuleDeclaration document)
                    {
                        if (document.IdentifierName.Equals("ThisWorkbook", StringComparison.InvariantCultureIgnoreCase))
                        {
                            document.AddSupertypeName("Workbook");
                            document.AddSupertypeName("_Workbook");
                        }
                        else if (document.IdentifierName.ToLowerInvariant().Contains("sheet"))
                        {
                            document.AddSupertypeName("Worksheet");
                            document.AddSupertypeName("_Worksheet");
                        }
                    }
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
