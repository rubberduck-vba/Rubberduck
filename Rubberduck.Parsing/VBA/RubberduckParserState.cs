using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Parsing.VBA
{
    public enum ResolutionState
    {
        Unresolved
    }

    public class ParserStateEventArgs : EventArgs
    {
        private readonly ParserState _state;

        public ParserStateEventArgs(ParserState state)
        {
            _state = state;
        }

        public ParserState State { get {return _state; } }
    }

    public sealed class RubberduckParserState
    {
        public event EventHandler ParseRequest;

        // keys are the declarations; values indicate whether a declaration is resolved.
        private readonly ConcurrentDictionary<Declaration, ResolutionState> _declarations =
            new ConcurrentDictionary<Declaration, ResolutionState>();

        private readonly ConcurrentDictionary<VBComponent, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<VBComponent, ITokenStream>();

        private readonly ConcurrentDictionary<VBComponent, IParseTree> _parseTrees =
            new ConcurrentDictionary<VBComponent, IParseTree>();

        public event EventHandler<ParserStateEventArgs> StateChanged;

        private void OnStateChanged(ParserState state)
        {
            var handler = StateChanged;
            if (handler != null)
            {
                handler.Invoke(this, new ParserStateEventArgs(state));
            }
        }

        private readonly ConcurrentDictionary<VBComponent, ParserState> _moduleStates =
            new ConcurrentDictionary<VBComponent, ParserState>();

        private readonly ConcurrentDictionary<VBComponent, SyntaxErrorException> _moduleExceptions =
            new ConcurrentDictionary<VBComponent, SyntaxErrorException>();

        private readonly object _lock = new object();
        public void SetModuleState(VBComponent component, ParserState state, SyntaxErrorException parserError = null)
        {
            _moduleStates[component] = state;
            _moduleExceptions[component] = parserError;

            // prevent multiple threads from changing state simultaneously:
            lock(_lock)
            {
                Status = EvaluateParserState();

            }
        }

        private ParserState EvaluateParserState()
        {
            var moduleStates = _moduleStates.Values.ToList();
            var state = Enum.GetValues(typeof (ParserState)).Cast<ParserState>()
                .SingleOrDefault(value => moduleStates.All(ps => ps == value));

            if (state != default(ParserState))
            {
                // if all modules are in the same state, we have our result.
                return state;
            }

            // intermediate states are toggled when *any* module has them.
            if (moduleStates.Any(ms => ms == ParserState.Error))
            {
                // error state takes precedence over every other state
                return ParserState.Error;
            }
            if (moduleStates.Any(ms => ms == ParserState.Parsing || ms == ParserState.Parsed))
            {
                return ParserState.Parsing;
            }
            if (moduleStates.Any(ms => ms == ParserState.Resolving))
            {
                return ParserState.Resolving;
            }

            return ParserState.Pending;
        }

        public ParserState GetModuleState(VBComponent component)
        {
            return _moduleStates[component];
        }

        private ParserState _status;
        public ParserState Status 
        { 
            get { return _status; }
            private set
            {
                if (_status != value)
                {
                    _status = value; 
                    OnStateChanged(value);
                }
            } 
        }

        private IEnumerable<QualifiedContext> _obsoleteCallContexts = new List<QualifiedContext>();

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing 'Call' statements in the parse tree.
        /// </summary>
        public IEnumerable<QualifiedContext> ObsoleteCallContexts
        {
            get { return _obsoleteCallContexts; }
            internal set { _obsoleteCallContexts = value; }
        }

        private IEnumerable<QualifiedContext> _obsoleteLetContexts = new List<QualifiedContext>();

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing explicit 'Let' statements in the parse tree.
        /// </summary>
        public IEnumerable<QualifiedContext> ObsoleteLetContexts
        {
            get { return _obsoleteLetContexts; }
            internal set { _obsoleteLetContexts = value; }
        }

        private IEnumerable<QualifiedContext> _emptyStringLiterals = new List<QualifiedContext>();

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing 'Call' statements in the parse tree.
        /// </summary>
        public IEnumerable<QualifiedContext> EmptyStringLiterals
        {
            get { return _emptyStringLiterals; }
            internal set { _emptyStringLiterals = value; }
        }

        private IEnumerable<QualifiedContext> _argListsWithOneByRefParam = new List<QualifiedContext>();

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing 'Call' statements in the parse tree.
        /// </summary>
        public IEnumerable<QualifiedContext> ArgListsWithOneByRefParam
        {
            get { return _argListsWithOneByRefParam; }
            internal set { _argListsWithOneByRefParam = value; }
        }

        private readonly ConcurrentDictionary<VBComponent, IEnumerable<CommentNode>> _comments =
            new ConcurrentDictionary<VBComponent, IEnumerable<CommentNode>>();

        public IEnumerable<CommentNode> Comments
        {
            get
            {
                return _comments.Values.SelectMany(comments => comments.ToList());
            }
        }

        public void SetModuleComments(VBComponent component, IEnumerable<CommentNode> comments)
        {
            _comments[component] = comments;
        }

        /// <summary>
        /// Gets a copy of the collected declarations.
        /// </summary>
        public IEnumerable<Declaration> AllDeclarations { get { return _declarations.Keys.ToList(); } }

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection (replaces existing).
        /// </summary>
        public void AddDeclaration(Declaration declaration)
        {
            if (_declarations.TryAdd(declaration, ResolutionState.Unresolved))
            {
                return;
            }

            if (RemoveDeclaration(declaration))
            {
                _declarations.TryAdd(declaration, ResolutionState.Unresolved);
            }
        }

        public void ClearDeclarations(VBComponent component)
        {
            var declarations = _declarations.Keys.Where(k =>
                k.QualifiedName.QualifiedModuleName.Project == component.Collection.Parent
                && k.ComponentName == component.Name);

            foreach (var declaration in declarations)
            {
                ResolutionState state;
                _declarations.TryRemove(declaration, out state);
            }
        }

        public void AddTokenStream(VBComponent component, ITokenStream stream)
        {
            _tokenStreams[component] = stream;
        }

        public void AddParseTree(VBComponent component, IParseTree parseTree)
        {
            _parseTrees[component] = parseTree;
        }

        public IEnumerable<KeyValuePair<VBComponent, IParseTree>> ParseTrees { get { return _parseTrees; } }

        public TokenStreamRewriter GetRewriter(VBComponent component)
        {
            return new TokenStreamRewriter(_tokenStreams[component]);
        }

        /// <summary>
        /// Removes the specified <see cref="declaration"/> from the collection.
        /// </summary>
        /// <param name="declaration"></param>
        /// <returns>Returns true when successful.</returns>
        private bool RemoveDeclaration(Declaration declaration)
        {
            ResolutionState state;
            return _declarations.TryRemove(declaration, out state);
        }

        /// <summary>
        /// Ensures parser state accounts for built-in declarations.
        /// This method has no effect if built-in declarations have already been loaded.
        /// </summary>
        public void AddBuiltInDeclarations(IHostApplication hostApplication)
        {
            if (_declarations.Any(declaration => declaration.Key.IsBuiltIn))
            {
                return;
            }

            var builtInDeclarations = VbaStandardLib.Declarations;

            // cannot be strongly-typed here because of constraints on COM interop and generics in the inheritance hierarchy. </rant>
            if (hostApplication /*is ExcelApp*/ .ApplicationName == "Excel")
            {
                builtInDeclarations = builtInDeclarations.Concat(ExcelObjectModel.Declarations);
            }

            foreach (var declaration in builtInDeclarations)
            {
                AddDeclaration(declaration);
            }
        }

        public void OnParseRequested()
        {
            var handler = ParseRequest;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }
    }
}