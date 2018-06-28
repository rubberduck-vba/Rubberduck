﻿using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Rubberduck.VBEditor;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols.ParsingExceptions;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ParseRunnerBase : IParseRunner
    {
        protected IParserStateManager StateManager { get; }

        private readonly RubberduckParserState _state;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly IAttributeParser _attributeParser;
        private readonly ISourceCodeHandler _sourceCodeHandler;

        protected ParseRunnerBase(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            Func<IVBAPreprocessor> preprocessorFactory,
            IAttributeParser attributeParser, 
            ISourceCodeHandler sourceCodeHandler)
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
            StateManager = parserStateManager;
            _preprocessorFactory = preprocessorFactory;
            _attributeParser = attributeParser;
            _sourceCodeHandler = sourceCodeHandler;
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
            var parser = new ComponentParseTask(module, preprocessor, _attributeParser, _sourceCodeHandler, _state.ProjectsProvider, rewriter);

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
                _state.SetModuleState(module, ParserState.Error, token, finishedParseTask.Exception?.InnerException as SyntaxErrorException);
            }
            else
            {
                var result = finishedParseTask.Result;
                lock (_state)
                {
                    token.ThrowIfCancellationRequested();

                    //This has to come first because it creates the module state if not present.
                    _state.SetModuleAttributes(module, result.Attributes);

                    _state.AddTokenStream(module, result.Tokens);
                    _state.AddParseTree(module, result.ParseTree);
                    _state.AddParseTree(module, result.AttributesTree, ParsePass.AttributesPass);
                    _state.SetModuleComments(module, result.Comments);
                    _state.SetModuleAnnotations(module, result.Annotations);
                    _state.AddAttributesRewriter(module, result.AttributesRewriter);

                    // This really needs to go last
                    //It does not reevaluate the overall parer state to avoid concurrent evaluation of all module states and for performance reasons.
                    //The evaluation has to be triggered manually in the calling procedure.
                    StateManager.SetModuleState(module, ParserState.Parsed, token, false); //Note that this is ok because locks allow re-entrancy.
                }
            }
        }
    }
}
