using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
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

        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<Declaration, byte>> _declarations =
            new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<Declaration, byte>>();

        private readonly ConcurrentDictionary<VBComponent, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<VBComponent, ITokenStream>();

        private readonly ConcurrentDictionary<VBComponent, IParseTree> _parseTrees =
            new ConcurrentDictionary<VBComponent, IParseTree>();

        public event EventHandler StateChanged;

        private void OnStateChanged()
        {
            var handler = StateChanged;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        private readonly ConcurrentDictionary<VBComponent, ParserState> _moduleStates =
            new ConcurrentDictionary<VBComponent, ParserState>();

        private readonly ConcurrentDictionary<VBComponent, SyntaxErrorException> _moduleExceptions =
            new ConcurrentDictionary<VBComponent, SyntaxErrorException>();

        public IReadOnlyList<Tuple<VBComponent, SyntaxErrorException>> ModuleExceptions
        {
            get { return _moduleExceptions.Select(kvp => Tuple.Create(kvp.Key, kvp.Value)).ToList(); }
        }

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
            _moduleStates.AddOrUpdate(component, state, (c, s) => state);
            _moduleExceptions.AddOrUpdate(component, parserError, (c, e) => parserError);

            Debug.WriteLine("Module '{0}' state is changing to '{1}' (thread {2})", component.Name, state, Thread.CurrentThread.ManagedThreadId);
            OnModuleStateChanged(component, state);
            Status = EvaluateParserState();
        }

        private static readonly ParserState[] States = Enum.GetValues(typeof (ParserState)).Cast<ParserState>().ToArray();

        private ParserState EvaluateParserState()
        {
            //lock (_lock)
            {
                var moduleStates = _moduleStates.Values.ToList();
                var state = States.SingleOrDefault(value => moduleStates.All(ps => ps == value));

                if (state != default(ParserState))
                {
                    // if all modules are in the same state, we have our result.
                    Debug.WriteLine("ParserState evaluates to '{0}' (thread {1})", state, Thread.CurrentThread.ManagedThreadId);
                    return state;
                }

                // error state takes precedence over every other state
                if (moduleStates.Any(ms => ms == ParserState.Error))
                {
                    Debug.WriteLine("ParserState evaluates to '{0}' (thread {1})", ParserState.Error,
                        Thread.CurrentThread.ManagedThreadId);
                    return ParserState.Error;
                }
                if (moduleStates.Any(ms => ms == ParserState.ResolverError))
                {
                    Debug.WriteLine("ParserState evaluates to '{0}' (thread {1})", ParserState.ResolverError,
                        Thread.CurrentThread.ManagedThreadId);
                    return ParserState.ResolverError;
                }

                // intermediate states are toggled when *any* module has them.
                var result = moduleStates.Min();
                if (moduleStates.Any(ms => ms == ParserState.Parsing))
                {
                    result = ParserState.Parsing;
                }
                if (moduleStates.Any(ms => ms == ParserState.Resolving))
                {
                    result = ParserState.Resolving;
                }

                Debug.WriteLine("ParserState evaluates to '{0}' (thread {1})", result,
                    Thread.CurrentThread.ManagedThreadId);
                return result;
            }
        }

        public ParserState GetModuleState(VBComponent component)
        {
            ParserState result;
            return _moduleStates.TryGetValue(component, out result) 
                ? result 
                : ParserState.Pending;
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
                    Debug.WriteLine("ParserState changed to '{0}', raising OnStateChanged", value);
                    OnStateChanged();
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
        public IReadOnlyList<Declaration> AllDeclarations 
        {
            get
            {
                return _declarations.Values.SelectMany(declarations => declarations.Keys).ToList();
            } 
        }

        /// <summary>
        /// Gets a copy of the collected declarations, excluding the built-in ones.
        /// </summary>
        public IReadOnlyList<Declaration> AllUserDeclarations
        {
            get
            {
                return _declarations.Values.Where(declarations => 
                        !declarations.Any(declaration => declaration.Key.IsBuiltIn))
                    .SelectMany(declarations => declarations.Keys)
                    .ToList();
            }
        }

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection (replaces existing).
        /// </summary>
        public void AddDeclaration(Declaration declaration)
        {
            var key = declaration.QualifiedName.QualifiedModuleName;
            var declarations = _declarations.GetOrAdd(key, new ConcurrentDictionary<Declaration, byte>());

            if (declarations.ContainsKey(declaration))
            {
                byte _;
                while (!declarations.TryRemove(declaration, out _))
                {
                    Debug.WriteLine("Could not remove existing declaration for '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
                }
            }
            while (!declarations.TryAdd(declaration, 0) && !declarations.ContainsKey(declaration))
            {
                Debug.WriteLine("Could not add declaration '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
            }
        }

        public void ClearDeclarations(VBProject project)
        {
            try
            {
                foreach (var component in project.VBComponents.Cast<VBComponent>())
                {
                    while (!ClearDeclarations(component))
                    {
                        // until Hell freezes over?
                    }
                }
            }
            catch (COMException)
            {
                _declarations.Clear();
            }
        }

        public void ClearBuiltInReferences()
        {
            foreach (var item in AllDeclarations.Where(item => item.IsBuiltIn && item.References.Any()))
            {
                item.ClearReferences();
            }
        }

        public bool ClearDeclarations(VBComponent component)
        {
            var project = component.Collection.Parent;
            var keys = _declarations.Keys.Where(kvp => 
                kvp.Project == project && kvp.Component == component);

            var success = true;
            foreach (var key in keys)
            {
                ConcurrentDictionary<Declaration, byte> declarations;
                success = success && (!_declarations.ContainsKey(key) || _declarations.TryRemove(key, out declarations));
            }
            
            ParserState state;
            success = success && (!_moduleStates.ContainsKey(component) || _moduleStates.TryRemove(component, out state));

            SyntaxErrorException exception;
            success = success && (!_moduleExceptions.ContainsKey(component) || _moduleExceptions.TryRemove(component, out exception));

            var components = _comments.Keys.Where(key =>
                key.Collection.Parent == project && key.Name == component.Name);

            foreach (var commentKey in components)
            {
                IList<CommentNode> nodes;
                success = success && (!_comments.ContainsKey(commentKey) || _comments.TryRemove(commentKey, out nodes));
            }

            Debug.WriteLine("ClearDeclarations({0}): {1}", component.Name, success ? "succeeded" : "failed");
            return success;
        }

        public void AddTokenStream(VBComponent component, ITokenStream stream)
        {
            _tokenStreams[component] = stream;
        }

        public void AddParseTree(VBComponent component, IParseTree parseTree)
        {
            _parseTrees[component] = parseTree;
        }

        public IParseTree GetParseTree(VBComponent component)
        {
            return _parseTrees[component];
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
            var key = declaration.QualifiedName.QualifiedModuleName;

            byte _;
            return _declarations[key].TryRemove(declaration, out _);
        }

        /// <summary>
        /// Ensures parser state accounts for built-in declarations.
        /// This method has no effect if built-in declarations have already been loaded.
        /// </summary>
        /// <summary>
        /// Requests reparse for specified component.
        /// Omit parameter to request a full reparse.
        /// </summary>
        /// <param name="requestor">The object requesting a reparse.</param>
        /// <param name="component">The component to reparse.</param>
        public void OnParseRequested(object requestor, VBComponent component = null)
        {
            var handler = ParseRequest;
            if (handler != null)
            {
                var args = new ParseRequestEventArgs(component);
                handler.Invoke(requestor, args);
            }
        }
    }
}