using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousParseRunner : ParseRunnerBase 
    {
        public SynchronousParseRunner(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            Func<IVBAPreprocessor> preprocessorFactory,
            IAttributeParser attributeParser,
            ISourceCodeHandler sourceCodeHandler,
            IModuleRewriterFactory moduleRewriterFactory) 
        :base(state, 
            parserStateManager, 
            preprocessorFactory, 
            attributeParser,
            sourceCodeHandler,
            moduleRewriterFactory)
        { }


        public override void ParseModules(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }

            token.ThrowIfCancellationRequested();

            try
            {
                foreach (var module in modules)
                {
                    ParseModule(module, token);
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception)
            {
                StateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                throw;
            }
        }
    }
}
