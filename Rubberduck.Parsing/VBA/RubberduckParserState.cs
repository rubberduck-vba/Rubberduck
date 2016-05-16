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

        public ParserState State { get { return _state; } }
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

    public sealed class RubberduckParserState : IDisposable
    {
        // circumvents VBIDE API's tendency to return a new instance at every parse, which breaks reference equality checks everywhere
        private readonly IDictionary<string, Func<VBProject>> _projects = new Dictionary<string, Func<VBProject>>();

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
                var args = new RubberduckStatusMessageEventArgs(message);
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
            if (_moduleStates.IsEmpty)
            {
                return ParserState.Pending;
            }

            var moduleStates = _moduleStates.Values.ToList();
            if (!moduleStates.Any())
            {
                return ParserState.Pending;
            }

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

            if (result == ParserState.Ready && moduleStates.Any(item => item != ParserState.Ready))
            {
                result = moduleStates.Except(new[] { ParserState.Ready }).Max();
            }

            Debug.Assert(result != ParserState.Ready || moduleStates.All(item => item == ParserState.Ready));

            Debug.WriteLine("ParserState evaluates to '{0}' (thread {1})", result,
            Thread.CurrentThread.ManagedThreadId);
            return result;
        }

        public ParserState GetOrCreateModuleState(VBComponent component)
        {
            var key = new QualifiedModuleName(component);
            var state = _moduleStates.GetOrAdd(key, ParserState.Pending);

            if (state == ParserState.Pending)
            {
                return state;   // we are slated for a reparse already
            }

            if (!IsNewOrModified(key))
            {
                return state;
            }

            _moduleStates.TryUpdate(key, ParserState.Pending, state);
            return ParserState.Pending;
        }

        public ParserState GetModuleState(VBComponent component)
        {
            return _moduleStates.GetOrAdd(new QualifiedModuleName(component), ParserState.Pending);
        }

        private ParserState _status;
        public ParserState Status
        {
            get { return _status; }
            internal set
            {
                if (_status != value)
                {
                    _status = value;
                    Debug.WriteLine("ParserState changed to '{0}', raising OnStateChanged", value);
                    OnStateChanged(_status);
                }
            }
        }

        internal void SetModuleAttributes(VBComponent component, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes)
        {
            _moduleAttributes.AddOrUpdate(new QualifiedModuleName(component), attributes, (c, s) => attributes);
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

        public void ClearStateCache(VBProject project, bool notifyStateChanged = false)
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

            if (notifyStateChanged)
            {
                OnStateChanged();
            }
        }

        public void ClearBuiltInReferences()
        {
            foreach (var item in AllDeclarations.Where(item => item.IsBuiltIn && item.References.Any()))
            {
                item.ClearReferences();
            }
        }

        public bool ClearStateCache(VBComponent component, bool notifyStateChanged = false)
        {
            var match = new QualifiedModuleName(component);
            var keys = _declarations.Keys.Where(kvp => kvp.Equals(match))
                .Union(new[] { match }).Distinct(); // make sure the key is present, even if there are no declarations left

            var success = RemoveKeysFromCollections(keys);

            var projectId = component.Collection.Parent.HelpFile;
            var sameProjectDeclarations = _declarations.Where(item => item.Key.ProjectId == projectId).ToList();
            if (sameProjectDeclarations.Any() &&
                sameProjectDeclarations.Count(item => item.Value.Any(key => key.Key.DeclarationType == DeclarationType.Project)) == sameProjectDeclarations.Count)
            {
                // only the project declaration is left; remove it.
                ConcurrentDictionary<Declaration, byte> declarations;
                _declarations.TryRemove(sameProjectDeclarations.Single().Key, out declarations);
                _projects.Remove(projectId);
                Debug.WriteLine(string.Format("Removed Project declaration for project Id {0}", projectId));
            }

            if (notifyStateChanged)
            {
                OnStateChanged();
            }

            return success;
        }

        public bool RemoveRenamedComponent(VBComponent component, string oldComponentName)
        {
            var match = new QualifiedModuleName(component, oldComponentName);
            var keys = _declarations.Keys.Where(kvp => kvp.ComponentName == oldComponentName && kvp.ProjectId == match.ProjectId);

            var success = RemoveKeysFromCollections(keys);

            OnStateChanged();
            return success;
        }

        private bool RemoveKeysFromCollections(IEnumerable<QualifiedModuleName> keys)
        {
            var success = true;
            foreach (var key in keys)
            {
                ConcurrentDictionary<Declaration, byte> declarations;
                success = success && (!_declarations.ContainsKey(key) || _declarations.TryRemove(key, out declarations));

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

        public bool IsDirty()
        {
            var projects = Projects.ToList();
            var components = projects.SelectMany(p => p.VBComponents.Cast<VBComponent>()).ToList();

            return components.Where(IsNewOrModified).Any();
        }

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
        private List<Tuple<Declaration, Selection, QualifiedModuleName>> _declarationSelections = new List<Tuple<Declaration, Selection, QualifiedModuleName>>();

        public void RebuildSelectionCache()
        {
            var declarations = AllDeclarations.Where(d => !d.IsBuiltIn).Select(d => Tuple.Create(d, d.Selection, d.QualifiedSelection.QualifiedName));
            var references = AllDeclarations.SelectMany(d => d.References.Select(r => Tuple.Create(d, r.Selection, r.QualifiedModuleName)));
            lock (_declarationSelections)
            {
                _declarationSelections.Clear();
                _declarationSelections.AddRange(declarations.Union(references));
            }
        }

        public Declaration FindSelectedDeclaration(CodePane activeCodePane, bool procedureLevelOnly = false)
        {

            var selection = activeCodePane.GetQualifiedSelection();
            if (selection.Equals(_lastSelection))
            {
                return _selectedDeclaration;
            }

            if (selection == null)
            {
                return _selectedDeclaration;
            }

            _lastSelection = selection.Value;
            _selectedDeclaration = null;

            if (!selection.Equals(default(QualifiedSelection)))
            {
                List<Tuple<Declaration, Selection, QualifiedModuleName>> matches = new List<Tuple<Declaration, Selection, QualifiedModuleName>>();
                lock (_declarationSelections)
                {
                    matches = _declarationSelections.Where(t =>
                                                    t.Item3.Equals(selection.Value.QualifiedName)
                                                    && (t.Item2.ContainsFirstCharacter(selection.Value.Selection))).ToList();
                }
                try
                {
                    if (matches.Count == 1)
                    {
                        _selectedDeclaration = matches.Single().Item1;
                    }
                    else
                    {
                        Declaration match = null;
                        if (procedureLevelOnly)
                        {
                            match = matches.Select(p => p.Item1).SingleOrDefault(item => item.DeclarationType.HasFlag(DeclarationType.Member));
                        }

                        // No match
                        if (matches.Count == 0)
                        {
                            match = match ?? AllUserDeclarations.SingleOrDefault(item =>
                                (item.DeclarationType == DeclarationType.ClassModule || item.DeclarationType == DeclarationType.ProceduralModule)
                                && item.QualifiedName.QualifiedModuleName.Equals(selection.Value.QualifiedName));
                        }
                        else
                        {
                            // Idiotic approach to find the best declaration out of a set of overlapping declarations.
                            // The one closest to the start of the user selection with the smallest width wins.
                            var userSelection = selection.Value.Selection;
                            var groupedByStartDistance = matches
                                .GroupBy(d => Tuple.Create(Math.Abs(userSelection.StartLine - d.Item2.StartLine), Math.Abs(userSelection.StartColumn - d.Item2.StartColumn)))
                                .OrderBy(g => g.Key.Item1)
                                .ThenBy(g => g.Key.Item2);
                            foreach (var closeMatch in groupedByStartDistance)
                            {
                                var groupedByLength = closeMatch.Select(d => Tuple.Create(d.Item1, Tuple.Create(Math.Abs(d.Item2.EndLine - d.Item2.StartLine), Math.Abs(d.Item2.EndColumn - d.Item2.StartColumn))))
                                    .OrderBy(d => d.Item2.Item1)
                                    .ThenBy(d => d.Item2.Item2).ToList();
                                match = groupedByLength.Select(p => p.Item1).FirstOrDefault();
                                break;
                            }
                        }

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

        public static Selection CreateBindingSelection(ParserRuleContext vbaGrammarContext, ParserRuleContext exprContext)
        {
            var k = exprContext.GetText();
            Selection vbaGrammarSelection = vbaGrammarContext.GetSelection();
            Selection exprSelection = exprContext.GetSelection();
            int lineOffset = vbaGrammarSelection.StartLine - 1;
            int columnOffset = 0;
            if (exprSelection.StartLine == 1)
            {
                columnOffset = vbaGrammarSelection.StartColumn - 1;
            }
            return new Selection(
                exprSelection.StartLine + lineOffset,
                exprSelection.StartColumn + columnOffset,
                exprSelection.EndLine + lineOffset,
                exprSelection.EndColumn + columnOffset);
        }

        public void RemoveBuiltInDeclarations(Reference reference)
        {
            var projectName = reference.Name;
            var key = new QualifiedModuleName(projectName, reference.FullPath, projectName);
            ConcurrentDictionary<Declaration, byte> items;
            if (!_declarations.TryRemove(key, out items))
            {
                Debug.WriteLine("Could not remove declarations for removed reference '{0}' ({1}).", reference.Name, QualifiedModuleName.GetProjectId(reference));
            }
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _declarations.Clear();
            _tokenStreams.Clear();
            _parseTrees.Clear();
            _moduleStates.Clear();
            _moduleContentHashCodes.Clear();
            _comments.Clear();
            _annotations.Clear();
            _moduleExceptions.Clear();
            _moduleAttributes.Clear();
            _declarationSelections.Clear();
            _projects.Clear();

            _isDisposed = true;
        }
    }
}
