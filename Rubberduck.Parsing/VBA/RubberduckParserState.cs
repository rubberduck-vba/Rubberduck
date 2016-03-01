using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

    public class ParseRequestEventArgs : EventArgs
    {
        private readonly VBComponent _component;

        public ParseRequestEventArgs(VBComponent component)
        {
            _component = component;
        }

        public VBComponent Component { get { return _component; } }
        public bool IsFullReparseRequest { get { return _component == null; } }
    }

    public sealed class RubberduckParserState
    {
        public event EventHandler<ParseRequestEventArgs> ParseRequest;

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

        public event EventHandler<ParseProgressEventArgs> ModuleStateChanged;

        private void OnModuleStateChanged(VBComponent component, ParserState state)
        {
            var handler = ModuleStateChanged;
            if (handler != null)
            {
                var args = new ParseProgressEventArgs(component, state);
                handler.Invoke(this, args);
            }
        }

        public void SetModuleState(VBComponent component, ParserState state, SyntaxErrorException parserError = null)
        {
            _moduleStates[component] = state;
            _moduleExceptions[component] = parserError;
            OnModuleStateChanged(component, state);
            Status = EvaluateParserState();
        }

        private static readonly ParserState[] States = Enum.GetValues(typeof (ParserState)).Cast<ParserState>().ToArray();

        private ParserState EvaluateParserState()
        {
            var moduleStates = _moduleStates.Values.ToList();
            var state = States.SingleOrDefault(value => moduleStates.All(ps => ps == value));

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
            if (moduleStates.Any(ms => ms == ParserState.Parsing))
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

        public IEnumerable<QualifiedContext> EmptyStringLiterals
        {
            get { return _emptyStringLiterals; }
            internal set { _emptyStringLiterals = value; }
        }

        private IEnumerable<QualifiedContext> _argListsWithOneByRefParam = new List<QualifiedContext>();

        public IEnumerable<QualifiedContext> ArgListsWithOneByRefParam
        {
            get { return _argListsWithOneByRefParam; }
            internal set { _argListsWithOneByRefParam = value; }
        }

        private readonly ConcurrentDictionary<VBComponent, IList<CommentNode>> _comments =
            new ConcurrentDictionary<VBComponent, IList<CommentNode>>();

        public IEnumerable<CommentNode> AllComments
        {
            get
            {
                return _comments.Values.SelectMany(comments => comments.ToList());
            }
        }

        public IEnumerable<CommentNode> GetModuleComments(VBComponent component)
        {
            IList<CommentNode> result;
            if (_comments.TryGetValue(component, out result))
            {
                return result;
            }

            return new List<CommentNode>();
        }

        public void SetModuleComments(VBComponent component, IEnumerable<CommentNode> comments)
        {
            _comments[component] = comments.ToList();
        }

        /// <summary>
        /// Gets a copy of the collected declarations, including the built-in ones.
        /// </summary>
        public IEnumerable<Declaration> AllDeclarations { get { return _declarations.Keys.ToList(); } }

        /// <summary>
        /// Gets a copy of the collected declarations, excluding the built-in ones.
        /// </summary>
        public IEnumerable<Declaration> AllUserDeclarations { get { return _declarations.Keys.Where(e => !e.IsBuiltIn).ToList(); } }

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

        public void ClearDeclarations(VBProject project)
        {
            try
            {
                foreach (var component in project.VBComponents.Cast<VBComponent>())
                {
                    ClearDeclarations(component);
                }
            }
            catch (COMException)
            {
                _declarations.Clear();
            }
        }

        public void ClearDeclarations(VBComponent component)
        {
            var declarations = _declarations.Keys.Where(k =>
                k.QualifiedName.QualifiedModuleName.Project == component.Collection.Parent
                && k.ComponentName == component.Name);

            foreach (var declaration in declarations)
            {
                RemoveDeclaration(declaration);
            }

            var components = _comments.Keys.Where(k =>
                k.Collection.Parent == component.Collection.Parent
                && k.Name == component.Name);

            foreach (var commentKey in components)
            {
                IList<CommentNode> nodes;
                _comments.TryRemove(commentKey, out nodes);
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
        public bool RemoveDeclaration(Declaration declaration)
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
            if (hostApplication != null && hostApplication.ApplicationName == "Excel")
            {
                builtInDeclarations = builtInDeclarations.Concat(ExcelObjectModel.Declarations);
            }

            foreach (var declaration in builtInDeclarations)
            {
                AddDeclaration(declaration);
            }
        }

        public void ResetBuiltInDeclarationReferences()
        {
            foreach (var item in _declarations.Keys.Where(declaration => declaration.IsBuiltIn))
            {
                item.ClearReferences();
            }
        }

        /// <summary>
        /// Requests reparse for specified component.
        /// Omit parameter to request a full reparse.
        /// </summary>
        /// <param name="component">The component to reparse.</param>
        public void OnParseRequested(VBComponent component = null)
        {
            var handler = ParseRequest;
            if (handler != null)
            {
                var args = new ParseRequestEventArgs(component);
                handler.Invoke(this, args);
            }
        }
    }
}