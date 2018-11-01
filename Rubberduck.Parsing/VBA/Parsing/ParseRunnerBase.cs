using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public abstract class ParseRunnerBase : IParseRunner
    {
        protected IParserStateManager StateManager { get; }

        private readonly RubberduckParserState _state;
        private readonly IModuleParser _parser;

        protected ParseRunnerBase(
            RubberduckParserState state,
            IParserStateManager parserStateManager,
            IModuleParser parser)
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

            _state = state;
            StateManager = parserStateManager;
            _parser = parser;
        }


        public abstract void ParseModules(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);


        protected void ParseModule(QualifiedModuleName module, CancellationToken token)
        {
            _state.ClearStateCache(module);
            try
            {
                var parseResults = _parser.Parse(module, token);
                SaveModuleParseResultsOnState(module, parseResults, token);
            }
            catch (SyntaxErrorException syntaxErrorException)
            {
                //In contrast to the situation in the success scenario, the overall parser state is reevaluated immediately.
                //This sets the state directly on the state because it is the sole instance where we have to pass the SyntaxErorException.
                _state.SetModuleState(module, ParserState.Error, token, syntaxErrorException);
            }
            catch (Exception exception)
            {
                StateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                throw;
            }
        }

        private void SaveModuleParseResultsOnState(QualifiedModuleName module, ModuleParseResults results, CancellationToken token)
        {
            lock (_state)
            {
                token.ThrowIfCancellationRequested();

                //This has to come first because it creates the module state if not present.
                _state.AddModuleStateIfNotPresent(module);

                _state.SaveContentHash(module);
                _state.AddParseTree(module, results.CodePaneParseTree);
                _state.AddParseTree(module, results.AttributesParseTree, CodeKind.AttributesCode);
                _state.SetModuleComments(module, results.Comments);
                _state.SetModuleAnnotations(module, results.Annotations);
                _state.SetModuleAttributes(module, results.Attributes);
                _state.SetMembersAllowingAttributes(module, results.MembersAllowingAttributes);
                _state.SetCodePaneTokenStream(module, results.CodePaneTokenStream);
                _state.SetAttributesTokenStream(module, results.AttributesTokenStream);

                // This really needs to go last
                //It does not reevaluate the overall parer state to avoid concurrent evaluation of all module states and for performance reasons.
                //The evaluation has to be triggered manually in the calling procedure.
                StateManager.SetModuleState(module, ParserState.Parsed, token, false); //Note that this is ok because locks allow re-entrancy.
            }
        }
    }
}
