using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public abstract class ParseRunnerBase : IParseRunner
    {
        protected IParserStateManager StateManager { get; }

        private readonly RubberduckParserState _state;
        private readonly IStringParser _parser;
        private readonly ISourceCodeProvider _codePaneSourceCodeProvider;
        private readonly ISourceCodeProvider _attributesSourceCodeProvider;
        private readonly IModuleRewriterFactory _moduleRewriterFactory;

        protected ParseRunnerBase(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IStringParser parser, 
            ISourceCodeProvider codePaneSourceCodeProvider,
            ISourceCodeProvider attributesSourceCodeProvider,
            IModuleRewriterFactory moduleRewriterFactory)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (parserStateManager == null)
            {
                throw new ArgumentNullException(nameof(parserStateManager));
            }
            if (parser == null)
            {
                throw new ArgumentNullException(nameof(parser));
            }
            if (moduleRewriterFactory == null)
            {
                throw new ArgumentNullException(nameof(moduleRewriterFactory));
            }
            if (codePaneSourceCodeProvider == null)
            {
                throw new ArgumentNullException(nameof(codePaneSourceCodeProvider));
            }
            if (attributesSourceCodeProvider == null)
            {
                throw new ArgumentNullException(nameof(attributesSourceCodeProvider));
            }

            _state = state;
            StateManager = parserStateManager;
            _parser = parser;
            _codePaneSourceCodeProvider = codePaneSourceCodeProvider;
            _attributesSourceCodeProvider = attributesSourceCodeProvider;
            _moduleRewriterFactory = moduleRewriterFactory;
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

            var parser = new ComponentParseTask(module, _codePaneSourceCodeProvider, _attributesSourceCodeProvider, _parser, _moduleRewriterFactory, rewriter);

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

                    _state.SaveContentHash(module);
                    _state.AddParseTree(module, result.ParseTree);
                    _state.AddParseTree(module, result.AttributesTree, CodeKind.AttributesCode);
                    _state.SetModuleComments(module, result.Comments);
                    _state.SetModuleAnnotations(module, result.Annotations);
                    _state.SetCodePaneRewriter(module, result.CodePaneRewriter);
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
