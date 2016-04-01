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

        private readonly ConcurrentDictionary<VBComponent, ParserState> _moduleStates =
            new ConcurrentDictionary<VBComponent, ParserState>();

        private readonly ConcurrentDictionary<VBComponent, IList<CommentNode>> _comments =
            new ConcurrentDictionary<VBComponent, IList<CommentNode>>();

        private readonly ConcurrentDictionary<VBComponent, SyntaxErrorException> _moduleExceptions =
            new ConcurrentDictionary<VBComponent, SyntaxErrorException>();

        private readonly ConcurrentDictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>> _moduleAttributes =
            new ConcurrentDictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>>();

        public IReadOnlyList<Tuple<VBComponent, SyntaxErrorException>> ModuleExceptions
        {
            get { return _moduleExceptions.Select(kvp => Tuple.Create(kvp.Key, kvp.Value)).Where(item => item.Item2 != null).ToList(); }
        }

        public event EventHandler<ParserStateEventArgs> StateChanged;

        private void OnStateChanged(ParserState state = ParserState.Pending)
        {
            var handler = StateChanged;
            if (handler != null)
            {
                handler.Invoke(this, new ParserStateEventArgs(state));
            }
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

        private ParserState EvaluateParserState()
        {
            var moduleStates = _moduleStates.Values.ToList();

            var prelim = moduleStates.Max();
            if (prelim == ParserState.Parsed && moduleStates.Any(s => s != ParserState.Parsed))
            {
                prelim = moduleStates.Where(s => s != ParserState.Parsed).Max();
            }
            return prelim;
        }

        public ParserState GetModuleState(VBComponent component)
        {
            return _moduleStates.GetOrAdd(component, ParserState.Pending);
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
                    OnStateChanged(_status);
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

        internal void SetModuleAttributes(VBComponent component, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes)
        {
            _moduleAttributes.AddOrUpdate(component, attributes, (c, s) => attributes);
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

        internal IDictionary<Tuple<string, DeclarationType>, Attributes> getModuleAttributes(VBComponent vbComponent)
        {
            return _moduleAttributes[vbComponent];
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
                kvp.Project == project && kvp.ComponentName == component.Name); // VBComponent reference seems to mismatch

            var success = true;
            var declarationsRemoved = 0;
            foreach (var key in keys)
            {
                ConcurrentDictionary<Declaration, byte> declarations = null;
                success = success && (!_declarations.ContainsKey(key) || _declarations.TryRemove(key, out declarations));
                declarationsRemoved = declarations == null ? 0 : declarations.Count;

                IParseTree tree;
                success = success && (!_parseTrees.ContainsKey(key.Component) || _parseTrees.TryRemove(key.Component, out tree));

                ITokenStream stream;
                success = success && (!_tokenStreams.ContainsKey(key.Component) || _tokenStreams.TryRemove(key.Component, out stream));

                ParserState state;
                success = success && (!_moduleStates.ContainsKey(key.Component) || _moduleStates.TryRemove(key.Component, out state));

                SyntaxErrorException exception;
                success = success && (!_moduleExceptions.ContainsKey(key.Component) || _moduleExceptions.TryRemove(key.Component, out exception));

                IList<CommentNode> nodes;
                success = success && (!_comments.ContainsKey(key.Component) || _comments.TryRemove(key.Component, out nodes));
            }

            Debug.WriteLine("ClearDeclarations({0}): {1} - {2} declarations removed", component.Name, success ? "succeeded" : "failed", declarationsRemoved);
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