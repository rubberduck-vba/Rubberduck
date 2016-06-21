using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using NLog;
// ReSharper disable LoopCanBeConvertedToQuery

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
        private readonly IDictionary<string, VBProject> _projects = new Dictionary<string, VBProject>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ModuleState> _moduleStates =
            new ConcurrentDictionary<QualifiedModuleName, ModuleState>();

        public event EventHandler<ParseRequestEventArgs> ParseRequest;
        public event EventHandler<RubberduckStatusMessageEventArgs> StatusMessageUpdate;

        private static readonly List<ParserState> States = new List<ParserState>();

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public readonly ConcurrentDictionary<List<string>, Declaration> CoClasses = new ConcurrentDictionary<List<string>, Declaration>();

        static RubberduckParserState()
        {
            var values = Enum.GetValues(typeof(ParserState));
            foreach (var value in values)
            {
                States.Add((ParserState)value);
            }
        }

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

            //assign a hashcode if no helpfile is present
            if (string.IsNullOrEmpty(project.HelpFile))
            {
                project.HelpFile = project.GetHashCode().ToString();
            }

            //loop until the helpfile is unique for this host session
            while (!IsProjectIdUnique(project.HelpFile))
            {
                project.HelpFile = (project.GetHashCode() ^ project.HelpFile.GetHashCode()).ToString();
            }

            var projectId = project.HelpFile;
            if (!_projects.ContainsKey(projectId))
            {
                _projects.Add(projectId, project);
            }

            foreach (VBComponent component in project.VBComponents)
            {
                _moduleStates.TryAdd(new QualifiedModuleName(component), new ModuleState(ParserState.Pending));
            }
        }

        private bool IsProjectIdUnique(string id)
        {
            foreach (var project in _projects)
            {
                if (project.Key == id)
                {
                    return false;
                }
            }

            return true;
        }

        public void RemoveProject(string projectId)
        {
            VBProject project = null;
            foreach (var p in Projects)
            {
                if (p.HelpFile == projectId)
                {
                    project = p;
                    break;
                }
            }

            if (_projects.ContainsKey(projectId))
            {
                _projects.Remove(projectId);
            }

            if (project != null)
            {
                ClearStateCache(project);
            }
        }

        public void RemoveProject(VBProject project)
        {
            RemoveProject(QualifiedModuleName.GetProjectId(project));
            ClearStateCache(project);
        }

        public List<VBProject> Projects
        {
            get
            {
                var projects = new List<VBProject>();
                foreach (var project in _projects.Values)
                {
                    projects.Add(project);
                }

                return projects;
            }
        }

        public IReadOnlyList<Tuple<VBComponent, SyntaxErrorException>> ModuleExceptions
        {
            get
            {
                var exceptions = new List<Tuple<VBComponent, SyntaxErrorException>>();
                foreach (var kvp in _moduleStates)
                {
                    if (kvp.Value.ModuleException == null)
                    {
                        continue;
                    }

                    exceptions.Add(Tuple.Create(kvp.Key.Component, kvp.Value.ModuleException));
                }

                return exceptions;
            }
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
            if (AllUserDeclarations.Count > 0)
            {
                var projectId = component.Collection.Parent.HelpFile;

                VBProject project = null;
                foreach (var item in _projects)
                {
                    if (item.Value.HelpFile == projectId)
                    {
                        project = project != null ? null : item.Value;
                    }
                }

                if (project == null)
                {
                    // ghost component shouldn't even exist
                    ClearStateCache(component);
                    Status = EvaluateParserState();
                    return;
                }
            }
            var key = new QualifiedModuleName(component);

            _moduleStates.AddOrUpdate(key, new ModuleState(state), (c, e) => e.SetState(state));
            _moduleStates.AddOrUpdate(key, new ModuleState(parserError), (c, e) => e.SetModuleException(parserError));
            Logger.Debug("Module '{0}' state is changing to '{1}' (thread {2})", key.ComponentName, state, Thread.CurrentThread.ManagedThreadId);
            OnModuleStateChanged(component, state);
            Status = EvaluateParserState();
        }

        private ParserState EvaluateParserState()
        {
            if (_moduleStates.IsEmpty)
            {
                return ParserState.Pending;
            }

            var moduleStates = new List<ParserState>();
            foreach (var moduleState in _moduleStates)
            {
                if (moduleState.Key.Component == null || moduleState.Key.ComponentName == string.Empty)
                {
                    continue;
                }

                moduleStates.Add(moduleState.Value.State);
            }

            if (moduleStates.Count == 0)
            {
                return ParserState.Pending;
            }

            var state = moduleStates[0];
            foreach (var moduleState in moduleStates)
            {
                if (moduleState != moduleStates[0])
                {
                    state = default(ParserState);
                    break;
                }
            }

            if (state != default(ParserState))
            {
                // if all modules are in the same state, we have our result.
                Logger.Debug("ParserState evaluates to '{0}' (thread {1})", state, Thread.CurrentThread.ManagedThreadId);
                return state;
            }

            var stateCounts = new int[States.Count];
            foreach (var moduleState in moduleStates)
            {
                stateCounts[(int)moduleState]++;
            }

            // error state takes precedence over every other state
            if (stateCounts[(int)ParserState.Error] > 0)
            {
                Logger.Debug("ParserState evaluates to '{0}' (thread {1})", ParserState.Error,
                Thread.CurrentThread.ManagedThreadId);
                return ParserState.Error;
            }
            if (stateCounts[(int)ParserState.ResolverError] > 0)
            {
                Logger.Debug("ParserState evaluates to '{0}' (thread {1})", ParserState.ResolverError,
                Thread.CurrentThread.ManagedThreadId);
                return ParserState.ResolverError;
            }

            // intermediate states are toggled when *any* module has them.
            var result = ParserState.None;
            foreach (var item in moduleStates)
            {
                if (item < result)
                {
                    result = item;
                }
            }

            if (stateCounts[(int)ParserState.Parsing] > 0)
            {
                result = ParserState.Parsing;
            }
            if (stateCounts[(int)ParserState.Resolving] > 0)
            {
                result = ParserState.Resolving;
            }

            if (result == ParserState.Ready)
            {
                for (var i = 0; i < stateCounts.Length; i++)
                {
                    if (i == (int)ParserState.Ready || i == (int)ParserState.None)
                    {
                        continue;
                    }

                    if (stateCounts[i] != 0)
                    {
                        result = (ParserState)i;
                    }
                }
            }

#if DEBUG
            if (state == ParserState.Ready)
            {
                for (var i = 0; i < stateCounts.Length; i++)
                {
                    if (i == (int)ParserState.Ready || i == (int)ParserState.None)
                    {
                        continue;
                    }

                    if (stateCounts[i] != 0)
                    {
                        Debug.Assert(false, "State is ready, but component has non-ready/non-none state");
                    }
                }
            }
#endif

            Logger.Debug("ParserState evaluates to '{0}' (thread {1})", result,
            Thread.CurrentThread.ManagedThreadId);
            return result;
        }

        public ParserState GetOrCreateModuleState(VBComponent component)
        {
            var key = new QualifiedModuleName(component);
            var state = _moduleStates.GetOrAdd(key, new ModuleState(ParserState.Pending)).State;

            if (state == ParserState.Pending)
            {
                return state;   // we are slated for a reparse already
            }

            if (!IsNewOrModified(key))
            {
                return state;
            }

            _moduleStates.AddOrUpdate(key, new ModuleState(ParserState.Pending), (c, s) => s.SetState(ParserState.Pending));
            return ParserState.Pending;
        }

        public ParserState GetModuleState(VBComponent component)
        {
            return _moduleStates.GetOrAdd(new QualifiedModuleName(component), new ModuleState(ParserState.Pending)).State;
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
                    Logger.Debug("ParserState changed to '{0}', raising OnStateChanged", value);
                    OnStateChanged(_status);
                }
            }
        }

        public void SetStatusAndFireStateChanged(ParserState status)
        {
            if (Status == status)
            {
                OnStateChanged(status);
            }
            else
            {
                Status = status;
            }
        }

        internal void SetModuleAttributes(VBComponent component, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes)
        {
            _moduleStates.AddOrUpdate(new QualifiedModuleName(component), new ModuleState(attributes), (c, s) => s.SetModuleAttributes(attributes));
        }

        public List<CommentNode> AllComments
        {
            get
            {
                var comments = new List<CommentNode>();
                foreach (var state in _moduleStates.Values)
                {
                    comments.AddRange(state.Comments);
                }

                return comments;
            }
        }

        public IEnumerable<CommentNode> GetModuleComments(VBComponent component)
        {
            ModuleState state;
            if (_moduleStates.TryGetValue(new QualifiedModuleName(component), out state))
            {
                return state.Comments;
            }

            return new List<CommentNode>();
        }

        public void SetModuleComments(VBComponent component, IEnumerable<CommentNode> comments)
        {
            _moduleStates[new QualifiedModuleName(component)].SetComments(new List<CommentNode>(comments));
        }

        public List<IAnnotation> AllAnnotations
        {
            get
            {
                var annotations = new List<IAnnotation>();
                foreach (var state in _moduleStates.Values)
                {
                    annotations.AddRange(state.Annotations);
                }

                return annotations;
            }
        }

        public IEnumerable<IAnnotation> GetModuleAnnotations(VBComponent component)
        {
            ModuleState result;
            if (_moduleStates.TryGetValue(new QualifiedModuleName(component), out result))
            {
                return result.Annotations;
            }

            return new List<IAnnotation>();
        }

        public void SetModuleAnnotations(VBComponent component, IEnumerable<IAnnotation> annotations)
        {
            _moduleStates[new QualifiedModuleName(component)].SetAnnotations(new List<IAnnotation>(annotations));
        }

        /// <summary>
        /// Gets a copy of the collected declarations, including the built-in ones.
        /// </summary>
        public IReadOnlyList<Declaration> AllDeclarations
        {
            get
            {
                var declarations = new List<Declaration>();
                foreach (var state in _moduleStates.Values)
                {
                    if (state.Declarations == null)
                    {
                        continue;
                    }

                    declarations.AddRange(state.Declarations.Keys);
                }

                return declarations;
            }
        }

        /// <summary>
        /// Gets a copy of the collected declarations, excluding the built-in ones.
        /// </summary>
        public IReadOnlyList<Declaration> AllUserDeclarations
        {
            get
            {
                var declarations = new List<Declaration>();
                foreach (var state in _moduleStates.Values)
                {
                    if (state.Declarations == null)
                    {
                        continue;
                    }

                    var hasBuiltInDeclaration = false;
                    foreach (var declaration in state.Declarations.Keys)
                    {
                        if (declaration.IsBuiltIn)
                        {
                            hasBuiltInDeclaration = true;
                            break;
                        }
                    }

                    if (!hasBuiltInDeclaration)
                    {
                        declarations.AddRange(state.Declarations.Keys);
                    }
                }

                return declarations;
            }
        }

        internal IDictionary<Tuple<string, DeclarationType>, Attributes> GetModuleAttributes(VBComponent vbComponent)
        {
            return _moduleStates[new QualifiedModuleName(vbComponent)].ModuleAttributes;
        }

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection (replaces existing).
        /// </summary>
        public void AddDeclaration(Declaration declaration)
        {
            var key = declaration.QualifiedName.QualifiedModuleName;
            var declarations = _moduleStates.GetOrAdd(key, new ModuleState(new ConcurrentDictionary<Declaration, byte>())).Declarations;

            if (declarations.ContainsKey(declaration))
            {
                byte _;
                while (!declarations.TryRemove(declaration, out _))
                {
                    Logger.Debug("Could not remove existing declaration for '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
                }
            }
            while (!declarations.TryAdd(declaration, 0) && !declarations.ContainsKey(declaration))
            {
                Logger.Debug("Could not add declaration '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
            }
        }

        public void ClearStateCache(VBProject project, bool notifyStateChanged = false)
        {
            try
            {
                foreach (VBComponent component in project.VBComponents)
                {
                    while (!ClearStateCache(component))
                    {
                        // until Hell freezes over?
                    }
                }
            }
            catch (COMException)
            {
                _moduleStates.Clear();
            }

            if (notifyStateChanged)
            {
                OnStateChanged(ParserState.ResolvedDeclarations);   // trigger test explorer and code explorer updates
                OnStateChanged(ParserState.Ready);   // trigger find all references &c. updates
            }
        }

        public void ClearBuiltInReferences()
        {
            foreach (var declaration in AllDeclarations)
            {
                if (!declaration.IsBuiltIn)
                {
                    continue;
                }

                var hasReference = false;
                foreach (var reference in declaration.References)
                {
                    hasReference = true;
                    break;
                }

                if (hasReference)
                {
                    declaration.ClearReferences();
                }
            }
        }

        public bool ClearStateCache(VBComponent component, bool notifyStateChanged = false)
        {
            if (component == null) { return false; }

            var match = new QualifiedModuleName(component);

            var keys = new List<QualifiedModuleName> { match };
            foreach (var key in _moduleStates.Keys)
            {
                if (key.Equals(match) && !keys.Contains(key))
                {
                    keys.Add(key);
                }
            }

            var success = RemoveKeysFromCollections(keys);

            var projectId = component.Collection.Parent.HelpFile;
            var sameProjectDeclarations = new List<KeyValuePair<QualifiedModuleName, ModuleState>>();
            foreach (var item in _moduleStates)
            {
                if (item.Key.ProjectId == projectId)
                {
                    sameProjectDeclarations.Add(new KeyValuePair<QualifiedModuleName, ModuleState>(item.Key, item.Value));
                }
            }

            var projectCount = 0;
            foreach (var item in sameProjectDeclarations)
            {
                if (item.Value.Declarations == null) { continue; }

                foreach (var declaration in item.Value.Declarations)
                {
                    if (declaration.Key.DeclarationType == DeclarationType.Project)
                    {
                        projectCount++;
                        break;
                    }
                }
            }

            if (sameProjectDeclarations.Count > 0 &&
                projectCount == sameProjectDeclarations.Count)
            {
                // only the project declaration is left; remove it.
                if (sameProjectDeclarations.Count != 1)
                {
                    throw new InvalidOperationException("Collection contains more than one item");
                }

                ModuleState moduleState;
                _moduleStates.TryRemove(sameProjectDeclarations[0].Key, out moduleState);
                if (moduleState != null)
                {
                    moduleState.Dispose();
                }

                _projects.Remove(projectId);
                Logger.Debug("Removed Project declaration for project Id {0}", projectId);
            }

            if (notifyStateChanged)
            {
                OnStateChanged(ParserState.ResolvedDeclarations);   // trigger test explorer and code explorer updates
                OnStateChanged(ParserState.Ready);   // trigger find all references &c. updates
            }

            return success;
        }

        public bool RemoveRenamedComponent(VBComponent component, string oldComponentName)
        {
            var match = new QualifiedModuleName(component, oldComponentName);
            var keys = new List<QualifiedModuleName>();
            foreach (var key in _moduleStates.Keys)
            {
                if (key.ComponentName == oldComponentName && key.ProjectId == match.ProjectId)
                {
                    keys.Add(key);
                }
            }

            var success = keys.Count != 0 && RemoveKeysFromCollections(keys);

            if (success)
            {
                OnStateChanged(ParserState.ResolvedDeclarations);   // trigger test explorer and code explorer updates
                OnStateChanged(ParserState.Ready);   // trigger find all references &c. updates
            }
            return success;
        }

        private bool RemoveKeysFromCollections(IEnumerable<QualifiedModuleName> keys)
        {
            var success = true;
            foreach (var key in keys)
            {
                ModuleState moduleState = null;
                success = success && (!_moduleStates.ContainsKey(key) || _moduleStates.TryRemove(key, out moduleState));

                if (moduleState != null)
                {
                    moduleState.Dispose();
                }
            }

            return success;
        }

        public void AddTokenStream(VBComponent component, ITokenStream stream)
        {
            _moduleStates[new QualifiedModuleName(component)].SetTokenStream(stream);
        }

        public void AddParseTree(VBComponent component, IParseTree parseTree)
        {
            var key = new QualifiedModuleName(component);
            _moduleStates[key].SetParseTree(parseTree);
            _moduleStates[key].SetModuleContentHashCode(key.ContentHashCode);
        }

        public IParseTree GetParseTree(VBComponent component)
        {
            return _moduleStates[new QualifiedModuleName(component)].ParseTree;
        }

        public List<KeyValuePair<QualifiedModuleName, IParseTree>> ParseTrees
        {
            get
            {
                var parseTrees = new List<KeyValuePair<QualifiedModuleName, IParseTree>>();
                foreach (var state in _moduleStates)
                {
                    if (state.Value.ParseTree != null)
                    {
                        parseTrees.Add(new KeyValuePair<QualifiedModuleName, IParseTree>(state.Key,
                            state.Value.ParseTree));
                    }
                }

                return parseTrees;
            }
        }

        public bool IsDirty()
        {
            foreach (var project in Projects)
            {
                foreach (VBComponent component in project.VBComponents)
                {
                    if (IsNewOrModified(component))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public bool HasAllParseTrees(IReadOnlyList<VBComponent> expected)
        {
            var expectedModules = new List<QualifiedModuleName>();
            foreach (var component in expected)
            {
                expectedModules.Add(new QualifiedModuleName(component));
            }

            foreach (var key in _moduleStates.Keys)
            {
                if (key.Component == null || expectedModules.Contains(key))
                {
                    continue;
                }

                ClearStateCache(key.Component);
            }

            var parseTreeCount = 0;
            foreach (var state in _moduleStates)
            {
                if (state.Value.ParseTree != null)
                {
                    parseTreeCount++;
                }
            }

            return parseTreeCount == expected.Count;
        }

        public TokenStreamRewriter GetRewriter(VBComponent component)
        {
            return new TokenStreamRewriter(_moduleStates[new QualifiedModuleName(component)].TokenStream);
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
            return _moduleStates[key].Declarations.TryRemove(declaration, out _);
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
            ModuleState moduleState;
            if (_moduleStates.TryGetValue(key, out moduleState))
            {
                // existing/modified
                return moduleState.IsNew || key.ContentHashCode != moduleState.ModuleContentHashCode;
            }

            // new
            return true;
        }

        private QualifiedSelection _lastSelection;
        private Declaration _selectedDeclaration;
        private readonly List<Tuple<Declaration, Selection, QualifiedModuleName>> _declarationSelections = new List<Tuple<Declaration, Selection, QualifiedModuleName>>();

        public void RebuildSelectionCache()
        {
            var selections = new List<Tuple<Declaration, Selection, QualifiedModuleName>>();
            foreach (var declaration in AllUserDeclarations)
            {
                selections.Add(Tuple.Create(declaration, declaration.Selection,
                    declaration.QualifiedSelection.QualifiedName));
            }

            foreach (var declaration in AllDeclarations)
            {
                foreach (var reference in declaration.References)
                {
                    selections.Add(Tuple.Create(declaration, reference.Selection, reference.QualifiedModuleName));
                }
            }

            lock (_declarationSelections)
            {
                _declarationSelections.Clear();
                _declarationSelections.AddRange(selections);
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
                var matches = new List<Tuple<Declaration, Selection, QualifiedModuleName>>();
                lock (_declarationSelections)
                {
                    foreach (var item in _declarationSelections)
                    {
                        if (item.Item3.Equals(selection.Value.QualifiedName) &&
                            item.Item2.ContainsFirstCharacter(selection.Value.Selection))
                        {
                            matches.Add(item);
                        }
                    }
                }
                try
                {
                    if (matches.Count == 1)
                    {
                        _selectedDeclaration = matches[0].Item1;
                    }
                    else
                    {
                        Declaration match = null;
                        if (procedureLevelOnly)
                        {
                            foreach (var item in matches)
                            {
                                if (item.Item1.DeclarationType.HasFlag(DeclarationType.Member))
                                {
                                    match = match != null ? null : item.Item1;
                                }
                            }
                        }

                        // No match
                        if (matches.Count == 0)
                        {
                            if (match == null)
                            {
                                foreach (var item in AllUserDeclarations)
                                {
                                    if ((item.DeclarationType == DeclarationType.ClassModule ||
                                         item.DeclarationType == DeclarationType.ProceduralModule) &&
                                        item.QualifiedName.QualifiedModuleName.Equals(selection.Value.QualifiedName))
                                    {
                                        match = match != null ? null : item;
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Idiotic approach to find the best declaration out of a set of overlapping declarations.
                            // The one closest to the start of the user selection with the smallest width wins.
                            var userSelection = selection.Value.Selection;

                            var currentSelection = matches[0].Item2;
                            match = matches[0].Item1;

                            foreach (var item in matches)
                            {
                                var itemDifferenceInStart = Math.Abs(userSelection.StartLine - item.Item2.StartLine);
                                var currentSelectionDifferenceInStart = Math.Abs(userSelection.StartLine - currentSelection.StartLine);

                                if (itemDifferenceInStart < currentSelectionDifferenceInStart)
                                {
                                    currentSelection = item.Item2;
                                    match = item.Item1;
                                }

                                if (itemDifferenceInStart == currentSelectionDifferenceInStart)
                                {
                                    if (Math.Abs(userSelection.StartColumn - item.Item2.StartColumn) <
                                        Math.Abs(userSelection.StartColumn - currentSelection.StartColumn))
                                    {
                                        currentSelection = item.Item2;
                                        match = item.Item1;
                                    }
                                }

                            }
                        }

                        _selectedDeclaration = match;
                    }
                }
                catch (InvalidOperationException exception)
                {
                    Logger.Error(exception);
                }
            }

            if (_selectedDeclaration != null)
            {
                Logger.Debug("Current selection ({0}) is '{1}' ({2})", selection, _selectedDeclaration.IdentifierName, _selectedDeclaration.DeclarationType);
            }

            return _selectedDeclaration;
        }

        public void RemoveBuiltInDeclarations(Reference reference)
        {
            var projectName = reference.Name;
            var key = new QualifiedModuleName(projectName, reference.FullPath, projectName);
            ModuleState moduleState;
            if (_moduleStates.TryRemove(key, out moduleState))
            {
                if (moduleState != null)
                {
                    moduleState.Dispose();
                }

                Logger.Warn("Could not remove declarations for removed reference '{0}' ({1}).", reference.Name, QualifiedModuleName.GetProjectId(reference));
            }
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            foreach (var item in _moduleStates)
            {
                item.Value.Dispose();
            }

            if (CoClasses != null)
            {
                CoClasses.Clear();
            }

            _moduleStates.Clear();
            _declarationSelections.Clear();
            _projects.Clear();

            _isDisposed = true;
        }
    }
}