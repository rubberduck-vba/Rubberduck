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
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

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

    public class RubberduckStatusMessageEventArgs : EventArgs
    {
        private readonly string _message;

        public RubberduckStatusMessageEventArgs(string message)
        {
            _message = message;
        }

        public string Message { get { return _message; } }
    }

    public sealed class RubberduckParserState
    {
        // circumvents VBIDE API's tendency to return a new instance at every parse, which breaks reference equality checks everywhere
        private readonly IDictionary<string,Func<VBProject>> _projects = new Dictionary<string,Func<VBProject>>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<Declaration, byte>> _declarations =
            new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<Declaration, byte>>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<QualifiedModuleName, ITokenStream>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IParseTree> _parseTrees =
            new ConcurrentDictionary<QualifiedModuleName, IParseTree>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ParserState> _moduleStates =
            new ConcurrentDictionary<QualifiedModuleName, ParserState>();

        private readonly ConcurrentDictionary<QualifiedModuleName, int> _moduleContentHashCodes =
            new ConcurrentDictionary<QualifiedModuleName, int>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IList<CommentNode>> _comments =
            new ConcurrentDictionary<QualifiedModuleName, IList<CommentNode>>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IList<IAnnotation>> _annotations =
            new ConcurrentDictionary<QualifiedModuleName, IList<IAnnotation>>();
        
        private readonly ConcurrentDictionary<QualifiedModuleName, SyntaxErrorException> _moduleExceptions =
            new ConcurrentDictionary<QualifiedModuleName, SyntaxErrorException>();

        private readonly ConcurrentDictionary<QualifiedModuleName, IDictionary<Tuple<string, DeclarationType>, Attributes>> _moduleAttributes =
            new ConcurrentDictionary<QualifiedModuleName, IDictionary<Tuple<string, DeclarationType>, Attributes>>();

        public event EventHandler<ParseRequestEventArgs> ParseRequest;
        public event EventHandler<RubberduckStatusMessageEventArgs> StatusMessageUpdate;

        public void OnStatusMessageUpdate(string message)
        {
            var handler = StatusMessageUpdate;
            if (handler != null)
            {
                var args=  new RubberduckStatusMessageEventArgs(message);
                handler.Invoke(this, args);
            }
        }

        public void AddProject(VBProject project)
        {
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                // adding protected project to parser state is asking for COMExceptions..
                return;
            }

            if (string.IsNullOrEmpty(project.HelpFile))
            {
                project.HelpFile = project.GetHashCode().ToString();
            }
            var projectId = project.HelpFile;
            if (!_projects.ContainsKey(projectId))
            {
                _projects.Add(projectId, () => project);
            }

            foreach (var component in project.VBComponents.Cast<VBComponent>())
            {
                _moduleStates.TryAdd(new QualifiedModuleName(component), ParserState.Pending);
            }
        }

        public void RemoveProject(string projectId)
        {
            if (_projects.ContainsKey(projectId))
            {
                _projects.Remove(projectId);
            }
        }

        public void RemoveProject(VBProject project)
        {
            RemoveProject(QualifiedModuleName.GetProjectId(project));
            ClearStateCache(project);
        }

        public IEnumerable<VBProject> Projects
        {
            get
            {
                return _projects.Values.Select(project => project.Invoke());
            }
        }

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

        public void SetModuleState(ParserState state)
        {
            var projects = Projects
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .ToList();

            var components = projects.SelectMany(p => p.VBComponents.Cast<VBComponent>()).ToList();
            foreach (var component in components)
            {
                SetModuleState(component, state);
            }
        }

        public void SetModuleState(VBComponent component, ParserState state, SyntaxErrorException parserError = null)
        {
            if (AllUserDeclarations.Any())
            {
                var projectId = component.Collection.Parent.HelpFile;
                var project = AllUserDeclarations.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.Project && item.ProjectId == projectId);

                if (project == null)
                {
                    // ghost component shouldn't even exist
                    ClearStateCache(component);
                    Status = EvaluateParserState();
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
            if (moduleStates.Count == 0)
            {
                return ParserState.Pending;
            }

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
                // now any module not ready means at least one of them has work in progress;
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

        public void ClearStateCache(VBProject project)
        {
            try
            {
                foreach (var component in project.VBComponents.Cast<VBComponent>())
                {
                    while (!ClearStateCache(component))
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

        public bool ClearStateCache(VBComponent component)
        {
            var match = new QualifiedModuleName(component);
            var keys = _declarations.Keys.Where(kvp => kvp.Equals(match))
                .Union(new[]{match}).Distinct(); // make sure the key is present, even if there are no declarations left

            var success = true;
            var declarationsRemoved = 0;
            foreach (var key in keys)
            {
                ConcurrentDictionary<Declaration, byte> declarations = null;
                success = success && (!_declarations.ContainsKey(key) || _declarations.TryRemove(key, out declarations));
                declarationsRemoved = declarations == null ? 0 : declarations.Count;

                IParseTree tree;
                success = success && (!_parseTrees.ContainsKey(key) || _parseTrees.TryRemove(key, out tree));

                int contentHash;
                success = success && (!_moduleContentHashCodes.ContainsKey(key) || _moduleContentHashCodes.TryRemove(key, out contentHash));

                IList<IAnnotation> annotations;
                success = success && (!_annotations.ContainsKey(key) || _annotations.TryRemove(key, out annotations));

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
            var key = new QualifiedModuleName(component);
            _parseTrees[key] = parseTree;
            _moduleContentHashCodes[key] = key.ContentHashCode;
        }

        public IParseTree GetParseTree(VBComponent component)
        {
            return _parseTrees[new QualifiedModuleName(component)];
        }

        public IEnumerable<KeyValuePair<QualifiedModuleName, IParseTree>> ParseTrees { get { return _parseTrees; } }

        public bool HasAllParseTrees(IReadOnlyList<VBComponent> expected)
        {
            var expectedModules = expected.Select(module => new QualifiedModuleName(module));
            foreach (var module in _moduleStates.Keys.Where(item => !expectedModules.Contains(item)))
            {
                ClearStateCache(module.Component);
            }

            return _parseTrees.Count == expected.Count;
        }        

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

        public bool IsNewOrModified(VBComponent component)
        {
            var key = new QualifiedModuleName(component);
            return IsNewOrModified(key);
        }
        
        public bool IsNewOrModified(QualifiedModuleName key)
        {
            int current;
            if (_moduleContentHashCodes.TryGetValue(key, out current))
            {
                // existing/modified
                return key.ContentHashCode != current;
            }

            // new
            return true;
        }

        private QualifiedSelection _lastSelection;
        private Declaration _selectedDeclaration;

        public Declaration FindSelectedDeclaration(CodePane activeCodePane)
        {
            var selection = activeCodePane.GetSelection();
            if (selection.Equals(_lastSelection))
            {
                return _selectedDeclaration;
            }

            _lastSelection = selection;
            _selectedDeclaration = null;

            if (!selection.Equals(default(QualifiedSelection)))
            {
                var matches = AllDeclarations
                    .Where(item => item.DeclarationType != DeclarationType.Project &&
                                   item.DeclarationType != DeclarationType.ModuleOption &&
                                   item.DeclarationType != DeclarationType.Class &&
                                   item.DeclarationType != DeclarationType.Module &&
                                   (IsSelectedDeclaration(selection, item) || item.References.Any(reference => reference.Declaration.Equals(item) && IsSelectedReference(selection, reference))))
                    .ToList();
                try
                {
                    if (matches.Count == 1)
                    {
                        _selectedDeclaration = matches.Single();
                    }
                    else
                    {
                        // ambiguous (?), or no match - make the module be the current selection
                        var match = AllUserDeclarations.SingleOrDefault(item =>
                                    (item.DeclarationType == DeclarationType.Class || item.DeclarationType == DeclarationType.Module)
                                    && item.QualifiedName.QualifiedModuleName.Equals(selection.QualifiedName));
                        _selectedDeclaration = match;
                    }
                }
                catch (InvalidOperationException exception)
                {
                    Debug.WriteLine(exception);
                }
            }

            if (_selectedDeclaration != null)
            {
                Debug.WriteLine("Current selection ({0}) is '{1}' ({2})", selection, _selectedDeclaration.IdentifierName, _selectedDeclaration.DeclarationType);
            }

            return _selectedDeclaration;
        }

        private static bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName)
                   && (declaration.QualifiedSelection.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private static bool IsSelectedReference(QualifiedSelection selection, IdentifierReference reference)
        {
            return reference.QualifiedModuleName.Equals(selection.QualifiedName)
                   && reference.Selection.ContainsFirstCharacter(selection.Selection);
        }

        public void RemoveBuiltInDeclarations(Reference reference)
        {
            var projectName = reference.Name;
            var path = reference.FullPath;
            var key = new QualifiedModuleName(projectName, path, projectName);
            ConcurrentDictionary<Declaration, byte> items;
            if (!_declarations.TryRemove(key, out items))
            {
                Debug.WriteLine("Could not remove declarations for removed reference '{0}' ({1}).", reference.Name, QualifiedModuleName.GetProjectId(reference));
            }
        }
    }
}