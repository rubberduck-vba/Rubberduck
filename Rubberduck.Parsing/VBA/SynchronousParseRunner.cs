using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousParseRunner : ParseRunnerBase 
    {
        public SynchronousParseRunner(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            Func<IVBAPreprocessor> preprocessorFactory,
            IAttributeParser attributeParser) 
        :base(state, 
            parserStateManager, 
            preprocessorFactory, 
            attributeParser)
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
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                throw;
            }
        }
    }
}
