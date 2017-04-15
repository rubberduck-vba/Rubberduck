using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using NLog;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

// ReSharper disable LoopCanBeConvertedToQuery

namespace Rubberduck.Parsing.VBA
{
    public class ParserStateEventArgs : EventArgs
    {
        public ParserStateEventArgs(ParserState state)
        {
            State = state;
        }

        public ParserState State { get; }
    }

    public class RubberduckStatusMessageEventArgs : EventArgs
    {
        public RubberduckStatusMessageEventArgs(string message)
        {
            Message = message;
        }

        public string Message { get; }
    }

    public sealed class RubberduckParserState : IDisposable, IDeclarationFinderProvider, IParseTreeProvider
    {
        // circumvents VBIDE API's tendency to return a new instance at every parse, which breaks reference equality checks everywhere
        private readonly IDictionary<string, IVBProject> _projects = new Dictionary<string, IVBProject>();

        private readonly ConcurrentDictionary<QualifiedModuleName, ModuleState> _moduleStates =
            new ConcurrentDictionary<QualifiedModuleName, ModuleState>();

        public event EventHandler<EventArgs> ParseRequest;
        public event EventHandler<RubberduckStatusMessageEventArgs> StatusMessageUpdate;

        private static readonly List<ParserState> States = new List<ParserState>();

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public readonly ConcurrentDictionary<List<string>, Declaration> CoClasses = new ConcurrentDictionary<List<string>, Declaration>();

        public bool IsEnabled { get; internal set; }

        public DeclarationFinder DeclarationFinder { get; private set; }

        internal void RefreshFinder(IHostApplication host)
        {
            DeclarationFinder = new DeclarationFinder(AllDeclarations, AllAnnotations, AllUnresolvedMemberDeclarations, host);
        }

        public void RefreshDeclarationFinder()
        {
            RefreshFinder(_hostApp);
        }

        private readonly IVBE _vbe;
        private readonly IHostApplication _hostApp;

        public RubberduckParserState(IVBE vbe)
        {
            var values = Enum.GetValues(typeof(ParserState));
            foreach (var value in values)
            {
                States.Add((ParserState)value);
            }

            _vbe = vbe;
            _hostApp = _vbe.HostApplication();
            AddEventHandlers();
            IsEnabled = true;
            RefreshFinder(_hostApp);
        }

        private void HandleStateChanged(ParserState state)
        {
            if (state == ParserState.ResolvedDeclarations || state == ParserState.Ready)
            {
                RefreshUserDeclarationsList();
            }
        }

        #region Event Handling

        private void AddEventHandlers()
        {
            VBProjects.ProjectAdded += Sinks_ProjectAdded;
            VBProjects.ProjectRemoved += Sinks_ProjectRemoved;
            VBProjects.ProjectRenamed += Sinks_ProjectRenamed;
            VBComponents.ComponentAdded += Sinks_ComponentAdded;
            VBComponents.ComponentRemoved += Sinks_ComponentRemoved;
            VBComponents.ComponentRenamed += Sinks_ComponentRenamed;           
        }

        private void RemoveEventHandlers()
        {
            VBProjects.ProjectAdded += Sinks_ProjectAdded;
            VBProjects.ProjectRemoved += Sinks_ProjectRemoved;
            VBProjects.ProjectRenamed += Sinks_ProjectRenamed;
            VBComponents.ComponentAdded -= Sinks_ComponentAdded;
            VBComponents.ComponentRemoved -= Sinks_ComponentRemoved;
            VBComponents.ComponentRenamed -= Sinks_ComponentRenamed;
        }

        private void Sinks_ProjectAdded(object sender, ProjectEventArgs e)
        {
            if (!e.Project.VBE.IsInDesignMode) { return; }

            Logger.Debug("Project '{0}' was added.", e.ProjectId);
            RefreshProjects(_vbe); // note side-effect: assigns ProjectId/HelpFile
            OnParseRequested(sender);
        }

        private void Sinks_ProjectRemoved(object sender, ProjectEventArgs e)
        {
            if (!e.Project.VBE.IsInDesignMode) { return; }
            
            Debug.Assert(e.ProjectId != null);
            RemoveProject(e.ProjectId, true);
            OnParseRequested(sender);
        }

        private void Sinks_ProjectRenamed(object sender, ProjectRenamedEventArgs e)
        {
            if (!e.Project.VBE.IsInDesignMode) { return; }

            if (AllDeclarations.Count == 0)
            {
                return;
            }

            Logger.Debug("Project {0} was renamed.", e.ProjectId);

            RemoveProject(e.ProjectId);
            RefreshProjects(e.Project.VBE);

            OnParseRequested(sender);
        }

        private void Sinks_ComponentAdded(object sender, ComponentEventArgs e)
        {
            if (!e.Project.VBE.IsInDesignMode) { return; }

            if (AllDeclarations.Count == 0)
            {
                return;
            }

            Logger.Debug("Component '{0}' was added.", e.Component.Name);
            OnParseRequested(sender);
        }

        private void Sinks_ComponentRemoved(object sender, ComponentEventArgs e)
        {
            if (!e.Project.VBE.IsInDesignMode) { return; }

            if (AllDeclarations.Count == 0)
            {
                return;
            }

            Logger.Debug("Component '{0}' was removed.", e.Component.Name);
            OnParseRequested(sender);
        }

        private void Sinks_ComponentRenamed(object sender, ComponentRenamedEventArgs e)
        {
            if (!e.Project.VBE.IsInDesignMode) { return; }

            if (AllDeclarations.Count == 0)
            {
                return;
            }

            Logger.Debug("Component '{0}' was renamed to '{1}'.", e.OldName, e.Component.Name);

            var componentIsWorksheet = false;
            foreach (var declaration in AllUserDeclarations)
            {
                if (declaration.ProjectId == e.ProjectId &&
                    declaration.DeclarationType == DeclarationType.ClassModule &&
                    declaration.IdentifierName == e.OldName)
                {
                    foreach (var superType in ((ClassModuleDeclaration) declaration).Supertypes)
                    {
                        if (superType.IdentifierName == "Worksheet")
                        {
                            componentIsWorksheet = true;
                            break;
                        }
                    }

                    break;
                }
            }

            if (componentIsWorksheet)
            {
                RemoveProject(e.ProjectId);
                Logger.Debug("Project '{0}' was removed.", e.Component.Name);

                RefreshProjects(e.Project.VBE);
            }
            else
            {
                RemoveRenamedComponent(e.ProjectId, e.OldName);
            }

            OnParseRequested(sender);
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

        #endregion

        /// <summary>
        /// Refreshes our list of cached projects.
        /// Be sure to reparse after calling this in case there
        /// were projects with duplicate ID's to clear the old
        /// declarations referencing the project by the old ID.
        /// </summary>
        public void RefreshProjects(IVBE vbe)
        {
            lock (_projects)
            {
                _projects.Clear();
                foreach (var project in vbe.VBProjects)
                {
                    if (project.Protection == ProjectProtection.Locked)
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(project.ProjectId) || _projects.Keys.Contains(project.ProjectId))
                    {
                        project.AssignProjectId();
                    }

                    _projects.Add(project.ProjectId, project);
                }
            }
        }

        //Overload using the vbe instance injected via the constructor.
        private void RefreshProjects()
        {
            RefreshProjects(_vbe);
        }

        private void RemoveProject(string projectId, bool notifyStateChanged = false)
        {
            lock (_projects)
            {
                if (_projects.ContainsKey(projectId))
                {
                    _projects.Remove(projectId);
                }
            }

            ClearStateCache(projectId, notifyStateChanged);
        }

        public List<IVBProject> Projects
        {
            get
            {
                lock(_projects)
                {
                    return new List<IVBProject>(_projects.Values);
                }
            }
        }

        public IReadOnlyList<Tuple<IVBComponent, SyntaxErrorException>> ModuleExceptions
        {
            get
            {
                var exceptions = new List<Tuple<IVBComponent, SyntaxErrorException>>();
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

        private int _stateChangedInvocations = 0;
        private void OnStateChanged(object requestor, ParserState state = ParserState.Pending)
        {
            _stateChangedInvocations++;
            HandleStateChanged(state);

            Logger.Info($"{nameof(RubberduckParserState)} ({_stateChangedInvocations}) is invoking {nameof(StateChanged)} ({Status})");
            StateChanged?.Invoke(requestor, new ParserStateEventArgs(state));
        }
        public event EventHandler<ParseProgressEventArgs> ModuleStateChanged;

        //Never spawn new threads changing module states in the handler! This will cause deadlocks. 
        private void OnModuleStateChanged(IVBComponent component, ParserState state, ParserState oldState)
        {
            var handler = ModuleStateChanged;
            if (handler != null)
            {
                var args = new ParseProgressEventArgs(component, state, oldState);
                handler.Invoke(this, args);
            }
        }

        public void SetModuleState(QualifiedModuleName module, ParserState state, CancellationToken token, SyntaxErrorException parserError = null, bool evaluateOverallState = true)
        {
            if (!token.IsCancellationRequested)
            {
                SetModuleState(module, state, parserError, evaluateOverallState);
            }
        }

        public void SetModuleState(QualifiedModuleName module, ParserState state, SyntaxErrorException parserError = null, bool evaluateOverallState = true)
        {
            if (AllUserDeclarations.Count > 0)
            {
                var projectId = module.ProjectId;
                IVBProject project = GetProject(projectId);

                if (project == null)
                {
                    // ghost component shouldn't even exist
                    ClearStateCache(module);
                    EvaluateParserState();
                    return;
                }
            }

            var oldState = GetModuleState(module);

            _moduleStates.AddOrUpdate(module, new ModuleState(state), (c, e) => e.SetState(state));
            _moduleStates.AddOrUpdate(module, new ModuleState(parserError), (c, e) => e.SetModuleException(parserError));
            Logger.Debug("Module '{0}' state is changing to '{1}' (thread {2})", module.ComponentName, state, Thread.CurrentThread.ManagedThreadId);
            OnModuleStateChanged(module.Component, state, oldState);
            if (evaluateOverallState)
            {
                EvaluateParserState();
            }
        }

        private IVBProject GetProject(string projectId)
        {
            IVBProject project = null;
            lock (_projects)
            {
                foreach (var item in _projects)
                {
                    if (item.Value.HelpFile == projectId)
                    {
                        if (project != null)
                        {
                            // ghost project detected, abort project iteration
                            project = null;
                            break;
                        }
                        project = item.Value;
                    }
                }
            }

            return project;
        }

        public void EvaluateParserState()
        {
            lock (_statusLockObject) Status = OverallParserStateFromModuleStates();
        }

        private ParserState OverallParserStateFromModuleStates()
        {
            if (_moduleStates.IsEmpty)
            {
                return ParserState.Pending;
            }

            var moduleStates = new List<ParserState>();
            foreach (var moduleState in _moduleStates)
            {
                if (moduleState.Key.Component == null || string.IsNullOrEmpty(moduleState.Key.ComponentName))
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
                return ParserState.Error;
            }
            if (stateCounts[(int)ParserState.ResolverError] > 0)
            {
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

            if (stateCounts[(int)ParserState.Pending] > 0)
            {
                result = ParserState.Pending;
            }
            if (stateCounts[(int)ParserState.Parsing] > 0)
            {
                result = ParserState.Parsing;
            }
            if (stateCounts[(int)ParserState.ResolvingDeclarations] > 0)
            {
                result = ParserState.ResolvingDeclarations;
            }
            if (stateCounts[(int)ParserState.ResolvingReferences] > 0)
            {
                result = ParserState.ResolvingReferences;
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
            
            return result;
        }

        public ParserState GetOrCreateModuleState(QualifiedModuleName module)
        {
            var state = _moduleStates.GetOrAdd(module, new ModuleState(ParserState.Pending)).State;

            if (state == ParserState.Pending)
            {
                return state;   // we are slated for a reparse already
            }

            if (!IsNewOrModified(module))
            {
                return state;
            }

            _moduleStates.AddOrUpdate(module, new ModuleState(ParserState.Pending), (c, s) => s.SetState(ParserState.Pending));
            return ParserState.Pending;
        }

        public ParserState GetModuleState(QualifiedModuleName module)
        {
            return _moduleStates.GetOrAdd(module, new ModuleState(ParserState.Pending)).State;
        }

        private readonly object _statusLockObject = new object(); 
        private ParserState _status;
        public ParserState Status
        {
            get { return _status; }
            private set
            {
                if (_status != value)
                {
                    _status = value;
                    OnStateChanged(this, _status);
                }
            }
        }

        public void SetStatusAndFireStateChanged(object requestor, ParserState status)
        {
            if (Status == status)
            {
                OnStateChanged(requestor, status);
            }
            else
            {
                Status = status;
            }
        }

        internal void SetModuleAttributes(QualifiedModuleName module, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes)
        {
            _moduleStates.AddOrUpdate(module, new ModuleState(attributes), (c, s) => s.SetModuleAttributes(attributes));
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

        public void SetModuleComments(QualifiedModuleName module, IEnumerable<CommentNode> comments)
        {
            _moduleStates[module].SetComments(new List<CommentNode>(comments));
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

        public IEnumerable<IAnnotation> GetModuleAnnotations(QualifiedModuleName module)
        {
            ModuleState result;
            if (_moduleStates.TryGetValue(module, out result))
            {
                return result.Annotations;
            }

            return new List<IAnnotation>();
        }

        public void SetModuleAnnotations(QualifiedModuleName module, IEnumerable<IAnnotation> annotations)
        {
            _moduleStates[module].SetAnnotations(new List<IAnnotation>(annotations));
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
        /// Gets a copy of the unresolved member declarations.
        /// </summary>
        public IReadOnlyList<UnboundMemberDeclaration> AllUnresolvedMemberDeclarations
        {
            get
            {
                var declarations = new List<UnboundMemberDeclaration>();
                foreach (var state in _moduleStates.Values)
                {
                    if (state.UnresolvedMemberDeclarations == null)
                    {
                        continue;
                    }

                    declarations.AddRange(state.UnresolvedMemberDeclarations.Keys);
                }

                return declarations;
            }
        }

        private readonly ConcurrentBag<SerializableProject> _builtInDeclarationTrees = new ConcurrentBag<SerializableProject>();
        public IProducerConsumerCollection<SerializableProject> BuiltInDeclarationTrees { get { return _builtInDeclarationTrees; } }

        private IReadOnlyList<Declaration> _allUserDeclarations = new List<Declaration>();

        /// <summary>
        /// Gets a copy of the collected declarations, excluding the built-in ones.
        /// </summary>
        public IReadOnlyList<Declaration> AllUserDeclarations => _allUserDeclarations;

        private void RefreshUserDeclarationsList()
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
                    if (!declaration.IsUserDefined)
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

            _allUserDeclarations = declarations;
        }

        internal IDictionary<Tuple<string, DeclarationType>, Attributes> GetModuleAttributes(QualifiedModuleName module)
        {
            return _moduleStates[module].ModuleAttributes;
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
                    Logger.Warn("Could not remove existing declaration for '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
                }
            }
            while (!declarations.TryAdd(declaration, 0) && !declarations.ContainsKey(declaration))
            {
                Logger.Warn("Could not add declaration '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
            }
        }

        /// <summary>
        /// Adds the specified <see cref="UnboundMemberDeclaration"/> to the collection (replaces existing).
        /// </summary>
        public void AddUnresolvedMemberDeclaration(UnboundMemberDeclaration declaration)
        {
            var key = declaration.QualifiedName.QualifiedModuleName;
            var declarations = _moduleStates.GetOrAdd(key, new ModuleState(new ConcurrentDictionary<Declaration, byte>())).UnresolvedMemberDeclarations;

            if (declarations.ContainsKey(declaration))
            {
                byte _;
                while (!declarations.TryRemove(declaration, out _))
                {
                    Logger.Warn("Could not remove existing unresolved member declaration for '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
                }
            }
            while (!declarations.TryAdd(declaration, 0) && !declarations.ContainsKey(declaration))
            {
                Logger.Warn("Could not add unresolved member declaration '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
            }
        }

        private void ClearStateCache(string projectId, bool notifyStateChanged = false)
        {
            try
            {
                foreach (var moduleState in _moduleStates)
                {
                    if (moduleState.Key.ProjectId == projectId && moduleState.Key.Component != null)
                    {
                        while (!ClearStateCache(moduleState.Key.Component))
                        {
                            // until Hell freezes over?
                        }
                    }
                    else if (moduleState.Key.ProjectId == projectId && moduleState.Key.Component == null)
                    {
                        // store project module name
                        var qualifiedModuleName = moduleState.Key;
                        ModuleState state;
                        if (_moduleStates.TryRemove(qualifiedModuleName, out state))
                        {
                            state.Dispose();
                        }
                    }
                }
            }
            catch (COMException)
            {
                _moduleStates.Clear();
            }

            if (notifyStateChanged)
            {
                OnStateChanged(this, ParserState.ResolvedDeclarations);   // trigger test explorer and code explorer updates
                OnStateChanged(this, ParserState.Ready);   // trigger find all references &c. updates
            }
        }

        public void ClearBuiltInReferences()
        {
            foreach (var declaration in AllDeclarations)
            {
                if (declaration.IsUserDefined)
                {
                    continue;
                }
                declaration.ClearReferences();
            }
        }

        public void ClearAllReferences()
        {
            foreach (var declaration in AllDeclarations)
            {
                declaration.ClearReferences();
            }
        }

        public bool ClearStateCache(IVBComponent component, bool notifyStateChanged = false)
        {
            return component != null && ClearStateCache(new QualifiedModuleName(component), notifyStateChanged);
        }

        public bool ClearStateCache(QualifiedModuleName module, bool notifyStateChanged = false)
        {
            var keys = new List<QualifiedModuleName> { module };
            foreach (var key in _moduleStates.Keys)
            {
                if (key.Equals(module) && !keys.Contains(key))
                {
                    keys.Add(key);
                }
            }

            var success = RemoveKeysFromCollections(keys);

            if (notifyStateChanged)
            {
                OnStateChanged(this, ParserState.ResolvedDeclarations);   // trigger test explorer and code explorer updates
                OnStateChanged(this, ParserState.Ready);   // trigger find all references &c. updates
            }

            return success;
        }

        private bool RemoveRenamedComponent(string projectId, string oldComponentName)
        {
            var keys = new List<QualifiedModuleName>();
            foreach (var key in _moduleStates.Keys)
            {
                if (key.ComponentName == oldComponentName && key.ProjectId == projectId)
                {
                    keys.Add(key);
                }
            }

            var success = keys.Count != 0 && RemoveKeysFromCollections(keys);

            if (success)
            {
                OnStateChanged(this, ParserState.ResolvedDeclarations);   // trigger test explorer and code explorer updates
                OnStateChanged(this, ParserState.Ready);   // trigger find all references &c. updates
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

        public void AddTokenStream(QualifiedModuleName module, ITokenStream stream)
        {
            _moduleStates[module].SetTokenStream(stream);
        }

        public void AddParseTree(QualifiedModuleName module, IParseTree parseTree)
        {
            var key = module;
            _moduleStates[key].SetParseTree(parseTree);
            _moduleStates[key].SetModuleContentHashCode(key.ContentHashCode);
        }

        public IParseTree GetParseTree(QualifiedModuleName module)
        {
            return _moduleStates[module].ParseTree;
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
                try
                {
                    foreach (var component in project.VBComponents)
                    {
                        if (IsNewOrModified(component))
                        {
                            return true;
                        }
                    }
                }
                catch (COMException)
                {
                }
            }

            return false;
        }

        public IModuleRewriter GetRewriter(IVBComponent component)
        {
            var qualifiedModuleName = new QualifiedModuleName(component);
            return GetRewriter(qualifiedModuleName);
        }

        public IModuleRewriter GetRewriter(QualifiedModuleName qualifiedModuleName)
        {
            var rewriter = _moduleStates[qualifiedModuleName].Rewriter;
            return new ModuleRewriter(qualifiedModuleName.Component.CodeModule, rewriter);
        }

        public IModuleRewriter GetRewriter(Declaration declaration)
        {
            var qualifiedModuleName = declaration.QualifiedSelection.QualifiedName;
            return GetRewriter(qualifiedModuleName);
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
        public void OnParseRequested(object requestor, IVBComponent component = null)
        {
            var handler = ParseRequest;
            if (handler != null && IsEnabled)
            {
                var args = EventArgs.Empty;
                handler.Invoke(requestor, args);
            }
        }

        public bool IsNewOrModified(IVBComponent component)
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

        public Declaration FindSelectedDeclaration(ICodePane activeCodePane)
        {
            return DeclarationFinder?.FindSelectedDeclaration(activeCodePane);
        }

        public void RemoveBuiltInDeclarations(IReference reference)
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

            RemoveEventHandlers();

            _moduleStates.Clear();
            // no lock because nobody should try to update anything here
            _projects.Clear();

            _isDisposed = true;
        }
    }
}