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
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Parsing.Annotations;

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

        // circumvents VBIDE API's tendency to return a new instance at every parse, which breaks reference equality checks everywhere
        private readonly IDictionary<string,VBProject> _projects = new Dictionary<string,VBProject>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<Declaration, byte>> _declarations =
            new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<Declaration, byte>>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<QualifiedModuleName, ITokenStream>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IParseTree> _parseTrees =
            new ConcurrentDictionary<QualifiedModuleName, IParseTree>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ParserState> _moduleStates =
            new ConcurrentDictionary<QualifiedModuleName, ParserState>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IList<CommentNode>> _comments =
            new ConcurrentDictionary<QualifiedModuleName, IList<CommentNode>>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IList<IAnnotation>> _annotations =
            new ConcurrentDictionary<QualifiedModuleName, IList<IAnnotation>>();
        
        private readonly ConcurrentDictionary<QualifiedModuleName, SyntaxErrorException> _moduleExceptions =
            new ConcurrentDictionary<QualifiedModuleName, SyntaxErrorException>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IDictionary<Tuple<string, DeclarationType>, Attributes>> _moduleAttributes =
            new ConcurrentDictionary<QualifiedModuleName, IDictionary<Tuple<string, DeclarationType>, Attributes>>();

        public void AddProject(VBProject project)
        {
            var name = project.ProjectName();
            if (!_projects.ContainsKey(name))
            {
                _projects.Add(name, project);
            }
        }

        public void RemoveProject(VBProject project)
        {
            var name = project.ProjectName();
            if (_projects.ContainsKey(name))
            {
                _projects.Remove(name);
            }
        }

        public IReadOnlyList<VBProject> Projects { get { return _projects.Values.ToList(); } }

        public IReadOnlyList<Tuple<VBComponent, SyntaxErrorException>> ModuleExceptions
        {
            get { return _moduleExceptions.Select(kvp => Tuple.Create(kvp.Key.Component, kvp.Value)).Where(item => item.Item2 != null).ToList(); }
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
            if (AllUserDeclarations.Any())
            {
                var projectName = component.ProjectName();
                var project = AllUserDeclarations.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.Project && item.ProjectName == projectName);

                if (project == null)
                {
                    // ghost component shouldn't even exist
                    ClearDeclarations(component);
                    return;
                }
            }
            var key = new QualifiedModuleName(component);
            _moduleStates.AddOrUpdate(key, state, (c, s) => state);
            _moduleExceptions.AddOrUpdate(key, parserError, (c, e) => parserError);

            Debug.WriteLine("Module '{0}' state is changing to '{1}' (thread {2})", key.ComponentName, state, Thread.CurrentThread.ManagedThreadId);
            OnModuleStateChanged(component, state);

            Status = EvaluateParserState();
        }

        private static readonly ParserState[] States = Enum.GetValues(typeof(ParserState)).Cast<ParserState>().ToArray();
        private ParserState EvaluateParserState()
        {
            var moduleStates = _moduleStates.Values.ToList();
            if (States.Any(state => moduleStates.All(module => module == state)))
            {
                // all modules have the same state - we're done here:
                return moduleStates.First();
            }

            if (moduleStates.Any(module => module > ParserState.Ready)) // only states beyond "ready" are error states
            {
                // any error state seals the deal:
                return moduleStates.Max();
            }

            if (moduleStates.Any(module => module != ParserState.Ready))
            {
                // any module not ready means at least one of them has work in progress;
                // report the least advanced of them, except if that's 'Pending':
                return moduleStates.Except(new[]{ParserState.Pending}).Min();
            }

            return default(ParserState); // default value is 'Pending'.
        }

        public ParserState GetModuleState(VBComponent component)
        {
            return _moduleStates.GetOrAdd(new QualifiedModuleName(component), ParserState.Pending);
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
            _moduleAttributes.AddOrUpdate(new QualifiedModuleName(component), attributes, (c, s) => attributes);
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
            if (_comments.TryGetValue(new QualifiedModuleName(component), out result))
            {
                return result;
            }

            return new List<CommentNode>();
        }

        public void SetModuleComments(VBComponent component, IEnumerable<CommentNode> comments)
        {
            _comments[new QualifiedModuleName(component)] = comments.ToList();
        }

        public IEnumerable<IAnnotation> AllAnnotations
        {
            get
            {
                return _annotations.Values.SelectMany(annotation => annotation.ToList());
            }
        }

        public IEnumerable<IAnnotation> GetModuleAnnotations(VBComponent component)
        {
            IList<IAnnotation> result;
            if (_annotations.TryGetValue(new QualifiedModuleName(component), out result))
            {
                return result;
            }

            return new List<IAnnotation>();
        }

        public void SetModuleAnnotations(VBComponent component, IEnumerable<IAnnotation> annotations)
        {
            _annotations[new QualifiedModuleName(component)] = annotations.ToList();
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

        internal IDictionary<Tuple<string, DeclarationType>, Attributes> GetModuleAttributes(VBComponent vbComponent)
        {
            return _moduleAttributes[new QualifiedModuleName(vbComponent)];
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
            var match = new QualifiedModuleName(component);
            var keys = _declarations.Keys.Where(kvp => kvp.Equals(match)); 

            var success = true;
            var declarationsRemoved = 0;
            foreach (var key in keys)
            {
                ConcurrentDictionary<Declaration, byte> declarations = null;
                success = success && (!_declarations.ContainsKey(key) || _declarations.TryRemove(key, out declarations));
                declarationsRemoved = declarations == null ? 0 : declarations.Count;

                IParseTree tree;
                success = success && (!_parseTrees.ContainsKey(key) || _parseTrees.TryRemove(key, out tree));

                ITokenStream stream;
                success = success && (!_tokenStreams.ContainsKey(key) || _tokenStreams.TryRemove(key, out stream));

                ParserState state;
                success = success && (!_moduleStates.ContainsKey(key) || _moduleStates.TryRemove(key, out state));

                SyntaxErrorException exception;
                success = success && (!_moduleExceptions.ContainsKey(key) || _moduleExceptions.TryRemove(key, out exception));

                IList<CommentNode> nodes;
                success = success && (!_comments.ContainsKey(key) || _comments.TryRemove(key, out nodes));
            }

            Debug.WriteLine("ClearDeclarations({0}): {1} - {2} declarations removed", component.Name, success ? "succeeded" : "failed", declarationsRemoved);
            return success;
        }

        public void AddTokenStream(VBComponent component, ITokenStream stream)
        {
            _tokenStreams[new QualifiedModuleName(component)] = stream;
        }

        public void AddParseTree(VBComponent component, IParseTree parseTree)
        {
            _parseTrees[new QualifiedModuleName(component)] = parseTree;
        }

        public IParseTree GetParseTree(VBComponent component)
        {
            return _parseTrees[new QualifiedModuleName(component)];
        }

        public IEnumerable<KeyValuePair<QualifiedModuleName, IParseTree>> ParseTrees { get { return _parseTrees; } }

        public TokenStreamRewriter GetRewriter(VBComponent component)
        {
            return new TokenStreamRewriter(_tokenStreams[new QualifiedModuleName(component)]);
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