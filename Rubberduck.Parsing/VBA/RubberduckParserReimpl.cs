using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Globalization;
using Rubberduck.Parsing.Preprocessing;
using System.Diagnostics;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing.VBA
{
    class RubberduckParserReimpl : IRubberduckParser
    {
        public RubberduckParserState State
        {
            get
            {
                return _state;
            }
        }

        private readonly CancellationTokenSource _central = new CancellationTokenSource();
        private readonly Dictionary<VBComponent, Tuple<Task, CancellationTokenSource>> _currentTasks = new Dictionary<VBComponent, Tuple<Task, CancellationTokenSource>>();

        private readonly Dictionary<VBComponent, IParseTree> _parseTrees = new Dictionary<VBComponent, IParseTree>();
        private readonly Dictionary<QualifiedModuleName, Dictionary<Declaration, byte>> _declarations = new Dictionary<QualifiedModuleName, Dictionary<Declaration, byte>>();
        private readonly Dictionary<VBComponent, ITokenStream> _tokenStreams = new Dictionary<VBComponent, ITokenStream>();
        private readonly Dictionary<VBComponent, IList<CommentNode>> _comments = new Dictionary<VBComponent, IList<CommentNode>>();
        private readonly IDictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>> _componentAttributes
            = new Dictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>>();


        private readonly ReferencedDeclarationsCollector _comReflector;

        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAttributeParser _attributeParser;

        public RubberduckParserReimpl(VBE vbe, RubberduckParserState state, IAttributeParser attributeParser)
        {
            _vbe = vbe;
            _state = state;
            _attributeParser = attributeParser;

            _comReflector = new ReferencedDeclarationsCollector();

            state.ParseRequest += ReparseRequested;
            state.StateChanged += StateOnStateChanged;
        }

        private void StateOnStateChanged(object sender, EventArgs e)
        {
            var args = e as ParserStateEventArgs;
            if (args.State == ParserState.Parsed)
            {
                Task.Run(() => Resolve(_central.Token));
                // Resolving should be triggered.. not our job?
            }
        }

        private void ReparseRequested(object sender, EventArgs e)
        {
            var args = e as ParseRequestEventArgs;
            if (args.IsFullReparseRequest)
            {
                Cancel();
                ParseAll();
            }
            else
            {
                Cancel(args.Component);
                ParseAsync(args.Component, CancellationToken.None, _state.GetRewriter(args.Component));
            }
        }

        /// <summary>
        /// Starts parsing all components of all unprotected VBProjects associated with the VBE-Instance passed to the constructor of this parser instance.
        /// </summary>
        private void ParseAll()
        {
            var projects = _vbe.VBProjects
                .Cast<VBProject>()
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none);

            var components = projects.SelectMany(p => p.VBComponents.Cast<VBComponent>());
            // invalidation cleanup should go into ParseAsync?
            foreach (var invalidated in _componentAttributes.Keys.Except(components))
            {
                _componentAttributes.Remove(invalidated);
            }

            foreach (var vbComponent in components)
            {
                while (!_state.ClearDeclarations(vbComponent)) { }

                ParseAsync(vbComponent, CancellationToken.None, _state.GetRewriter(vbComponent));
            }
        }

        public Task ParseAsync(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            // FIXME remove invalidated "things"from _state
            // this includes: Declarations, Comments, Attributes, "InspectionResults" (ObsoleteCall, ObsoleteLet, EmptyStringLiteral, ArgLists with OneByRef) and possibly more...
            _state.ClearDeclarations(component);
            _state.SetModuleState(component, ParserState.Pending);
            _state.SetModuleComments(component, Enumerable.Empty<CommentNode>());
            _state.SetModuleAttributes(component, new Dictionary<Tuple<string, DeclarationType>, Attributes>());

            var linkedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token, token);

            var task = new Task(() => ParseAsyncInternal(component, linkedTokenSource.Token, rewriter));
            _currentTasks.Add(component, Tuple.Create(task, linkedTokenSource));

            task.ContinueWith(t => _currentTasks.Remove(component)); // default also executes on cancel

            _state.SetModuleState(component, ParserState.Parsing);
            task.Start();

            return task;
        }

        public void Cancel(VBComponent component = null)
        {            
            if (component == null)
            {
                _central.Cancel(false);
            }
            else
            {
                Tuple<Task, CancellationTokenSource> result;
                if (_currentTasks.TryGetValue(component, out result))
                {
                    result.Item2.Cancel(false);
                    // should we do this??
                    //result.Item1.Wait(); 
                }
            }
        }

        private void ParseAsyncInternal(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            var preprocessor = new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture));
            var parser = new ComponentParseTask(component, preprocessor, _attributeParser, _state.GetRewriter(component));
            parser.ParseFailure += (sender, e) => _state.SetModuleState(component, ParserState.Error, e.Cause as SyntaxErrorException);
            parser.ParseCompleted += (sender, e) =>
            {
                // possibly lock _state
                _state.SetModuleState(component, ParserState.Parsed);
                _state.SetModuleAttributes(component, e.Attributes);
                _state.AddParseTree(component, e.ParseTree);
                _state.AddTokenStream(component, e.Tokens);
                _state.SetModuleComments(component, e.Comments);
            };
            var task = parser.ParseAsync(token);
            _state.SetModuleState(component, ParserState.Parsing);
            task.Start();
        }

        public void ParseComponent(VBComponent component, TokenStreamRewriter rewriter = null)
        {
            ParseAsync(component, CancellationToken.None, rewriter).Wait();
        }

        public void Resolve(CancellationToken token)
        {
            Task.Run(() => ResolveInternal(token));
        }

        private void ResolveInternal(CancellationToken token)
        {
            foreach (var kvp in _state.ParseTrees)
            {
                if (token.IsCancellationRequested) return;
                ResolveDeclarations(kvp.Key, kvp.Value);
            }
            var finder = new DeclarationFinder(_state.AllDeclarations, _state.AllComments);
            foreach (var kvp in _state.ParseTrees)
            {
                if (token.IsCancellationRequested) return;
                ResolveReferences(finder, kvp.Key, kvp.Value);
            }
        }

        private void ResolveDeclarations(VBComponent component, IParseTree tree)
        {
            // cannot locate declarations in one pass *the way it's currently implemented*,
            // because the context in EnterSubStmt() doesn't *yet* have child nodes when the context enters.
            // so we need to EnterAmbiguousIdentifier() and evaluate the parent instead - this *might* work.
            var declarations = new List<Declaration>();
            var qualifiedModuleName = new QualifiedModuleName(component);
            DeclarationSymbolsListener declarationsListener = new DeclarationSymbolsListener(qualifiedModuleName, Accessibility.Implicit, component.Type, _state.GetModuleComments(component), _state.getModuleAttributes(component));
            declarationsListener.NewDeclaration += (sender, e) => _state.AddDeclaration(e.Declaration);
            declarationsListener.CreateModuleDeclarations();

            var obsoleteCallStatementListener = new ObsoleteCallStatementListener();
            var obsoleteLetStatementListener = new ObsoleteLetStatementListener();
            var emptyStringLiteralListener = new EmptyStringLiteralListener();
            var argListWithOneByRefParamListener = new ArgListWithOneByRefParamListener();

            var walker = new ParseTreeWalker();
            walker.Walk(new CombinedParseTreeListener(new IParseTreeListener[]{
                obsoleteCallStatementListener,
                obsoleteLetStatementListener,
                emptyStringLiteralListener,
                argListWithOneByRefParamListener,
                
                declarationsListener,
            }), tree);

            // FIXME this are actually (almost) isnpection results.. we should handle them as such
            _state.ArgListsWithOneByRefParam = argListWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
            _state.EmptyStringLiterals = emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
            _state.ObsoleteLetContexts = obsoleteLetStatementListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
            _state.ObsoleteCallContexts = obsoleteCallStatementListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
            
        }
        
        private void ResolveReferences(DeclarationFinder finder, VBComponent component, IParseTree tree)
        {
            var state = _state.GetModuleState(component);
            if (_state.Status == ParserState.ResolverError || state != ParserState.Parsed)
            {
                return;
            }
            _state.SetModuleState(component, ParserState.Resolving);
            Debug.WriteLine("Resolving '{0}'... (thread {1})", component.Name, Thread.CurrentThread.ManagedThreadId);

            var qualifiedName = new QualifiedModuleName(component);
            var resolver = new IdentifierReferenceResolver(qualifiedName, finder);
            var listener = new IdentifierReferenceListener(resolver);
            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var walker = new ParseTreeWalker();
                try
                {
                    walker.Walk(listener, tree);
                    state = ParserState.Ready;
                }
                catch (Exception exception)
                {
                    Debug.Print("Exception thrown resolving '{0}' (thread {2}): {1}", component.Name, exception, Thread.CurrentThread.ManagedThreadId);
                    state = ParserState.ResolverError;
                }
            }

            _state.SetModuleState(component, state);
            Debug.Print("'{0}' is {1}. Resolver took {2}ms to complete (thread {3})", component.Name, _state.GetModuleState(component), /*_resolverTimer[component].ElapsedMilliseconds*/0, Thread.CurrentThread.ManagedThreadId);
        }

        #region Listener classes
        private class ObsoleteCallStatementListener : VBABaseListener
        {
            private readonly IList<VBAParser.ExplicitCallStmtContext> _contexts = new List<VBAParser.ExplicitCallStmtContext>();
            public IEnumerable<VBAParser.ExplicitCallStmtContext> Contexts { get { return _contexts; } }

            public override void ExitExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
            {
                var procedureCall = context.eCS_ProcedureCall();
                if (procedureCall != null)
                {
                    if (procedureCall.CALL() != null)
                    {
                        _contexts.Add(context);
                        return;
                    }
                }

                var memberCall = context.eCS_MemberProcedureCall();
                if (memberCall == null) return;
                if (memberCall.CALL() == null) return;
                _contexts.Add(context);
            }
        }

        private class ObsoleteLetStatementListener : VBABaseListener
        {
            private readonly IList<VBAParser.LetStmtContext> _contexts = new List<VBAParser.LetStmtContext>();
            public IEnumerable<VBAParser.LetStmtContext> Contexts { get { return _contexts; } }

            public override void ExitLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                    _contexts.Add(context);
                }
            }
        }

        private class EmptyStringLiteralListener : VBABaseListener
        {
            private readonly IList<VBAParser.LiteralContext> _contexts = new List<VBAParser.LiteralContext>();
            public IEnumerable<VBAParser.LiteralContext> Contexts { get { return _contexts; } }

            public override void ExitLiteral(VBAParser.LiteralContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal != null && literal.GetText() == "\"\"")
                {
                    _contexts.Add(context);
                }
            }
        }

        private class ArgListWithOneByRefParamListener : VBABaseListener
        {
            private readonly IList<VBAParser.ArgListContext> _contexts = new List<VBAParser.ArgListContext>();
            public IEnumerable<VBAParser.ArgListContext> Contexts { get { return _contexts; } }

            public override void ExitArgList(VBAParser.ArgListContext context)
            {
                if (context.arg() != null && context.arg().Count(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null)) == 1)
                {
                    _contexts.Add(context);
                }
            }
        }

        #endregion
    }
}
