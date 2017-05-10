using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Rubberduck.VBEditor;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ParseRunnerBase : IParseRunner
    {
        private readonly RubberduckParserState _state;
        protected readonly IParserStateManager _parserStateManager;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly IAttributeParser _attributeParser;


        public ParseRunnerBase(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            Func<IVBAPreprocessor> preprocessorFactory,
            IAttributeParser attributeParser)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (parserStateManager == null)
            {
                throw new ArgumentNullException(nameof(parserStateManager));
            }
            if (preprocessorFactory == null)
            {
                throw new ArgumentNullException(nameof(preprocessorFactory));
            }
            if (attributeParser == null)
            {
                throw new ArgumentNullException(nameof(attributeParser));
            }

            _state = state;
            _parserStateManager = parserStateManager;
            _preprocessorFactory = preprocessorFactory;
            _attributeParser = attributeParser;
        }


        public abstract void ParseModules(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);


        protected void ParseModule(QualifiedModuleName module, CancellationToken token)
        {
            _state.ClearStateCache(module);
            var finishedParseTask = FinishedParseComponentTask(module, token);
            ProcessComponentParseResults(module, finishedParseTask, token);
        }

        private Task<ComponentParseTask.ParseCompletionArgs> FinishedParseComponentTask(QualifiedModuleName module, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            var tcs = new TaskCompletionSource<ComponentParseTask.ParseCompletionArgs>();

            var preprocessor = _preprocessorFactory();
            var parser = new ComponentParseTask(module, preprocessor, _attributeParser, rewriter);

            parser.ParseFailure += (sender, e) =>
            {
                if (e.Cause is OperationCanceledException)
                {
                    tcs.SetCanceled();
                }
                else
                {
                    tcs.SetException(e.Cause);
                }
            };
            parser.ParseCompleted += (sender, e) =>
            {
                tcs.SetResult(e);
            };

            parser.Start(token);

            return tcs.Task;
        }

        private void ProcessComponentParseResults(QualifiedModuleName module, Task<ComponentParseTask.ParseCompletionArgs> finishedParseTask, CancellationToken token)
        {
            if (finishedParseTask.IsFaulted)
            {
                //In contrast to the situation in the success scenario, the overall parser state is reevaluated immediately.
                //This sets the state directly on the state because it is the sole instance where we have to pass the potential SyntaxErorException.
                _state.SetModuleState(module, ParserState.Error, token, finishedParseTask.Exception.InnerException as SyntaxErrorException);
            }
            else
            {
                var result = finishedParseTask.Result;
                lock (_state)
                {
                    lock (module.Component)
                    {
                        _state.SetModuleAttributes(module, result.Attributes);
                        _state.AddParseTree(module, result.ParseTree);
                        _state.AddTokenStream(module, result.Tokens);
                        _state.SetModuleComments(module, result.Comments);
                        _state.SetModuleAnnotations(module, result.Annotations);

                        // This really needs to go last
                        //It does not reevaluate the overall parer state to avoid concurrent evaluation of all module states and for performance reasons.
                        //The evaluation has to be triggered manually in the calling procedure.
                        _parserStateManager.SetModuleState(module, ParserState.Parsed, token, false); //Note that this is ok because locks allow re-entrancy.
                    }
                }
            }
        }
    }
}
