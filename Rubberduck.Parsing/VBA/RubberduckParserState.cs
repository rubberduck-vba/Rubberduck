using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using NLog;
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
        private readonly ParserState _state;

        public ParserStateEventArgs(ParserState state)
        {
            _state = state;
        }

        public ParserState State { get { return _state; } }
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

        private readonly IVBE _vbe;
        public RubberduckParserState(IVBE vbe)
        {
            var values = Enum.GetValues(typeof(ParserState));
            foreach (var value in values)
            {
                States.Add((ParserState)value);
            }

            _vbe = vbe;
            AddEventHandlers();
            IsEnabled = true;
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

        private void OnStateChanged(object requestor, ParserState state = ParserState.Pending)
        {
            var handler = StateChanged;
            Logger.Debug("RubberduckParserState raised StateChanged ({0})", Status);
            if (handler != null)
            {               
                handler.Invoke(requestor, new ParserStateEventArgs(state));
            }
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


        public void SetModuleState(IVBComponent component, ParserState state, CancellationToken token, SyntaxErrorException parserError = null, bool evaluateOverallState = true)
        {
            if (!token.IsCancellationRequested)
            {
                SetModuleState(component, state, parserError, evaluateOverallState);
            }
        }
        
        public void SetModuleState(IVBComponent component, ParserState state, SyntaxErrorException parserError = null, bool evaluateOverallState = true)
        {
            if (AllUserDeclarations.Count > 0)
            {
                var projectId = component.Collection.Parent.HelpFile;

                IVBProject project = null;
                lock (_projects)
                {
                    foreach (var item in _projects)
                    {
                        if (item.Value.HelpFile == projectId)
                        {
                            if (project != null)
                            {
                                // ghost component detected, abort project iteration
                                project = null;
                                break;
                            }
                            project = item.Value;
                        }
                    }
                }

                if (project == null)
                {
                    // ghost component shouldn't even exist
                    ClearStateCache(component);
                    EvaluateParserState();
                    return;
                }
            }
            var key = new QualifiedModuleName(component);

            var oldState = GetModuleState(component);

            _moduleStates.AddOrUpdate(key, new ModuleState(state), (c, e) => e.SetState(state));
            _moduleStates.AddOrUpdate(key, new ModuleState(parserError), (c, e) => e.SetModuleException(parserError));
            Logger.Debug("Module '{0}' state is changing to '{1}' (thread {2})", key.ComponentName, state, Thread.CurrentThread.ManagedThreadId);
            OnModuleStateChanged(component, state, oldState);
            if (evaluateOverallState)
            {
                EvaluateParserState();
            }
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

        public ParserState GetOrCreateModuleState(IVBComponent component)
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

        public ParserState GetModuleState(IVBComponent component)
        {
            return _moduleStates.GetOrAdd(new QualifiedModuleName(component), new ModuleState(ParserState.Pending)).State;
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

        internal void SetModuleAttributes(IVBComponent component, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes)
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

        public void SetModuleComments(IVBComponent component, IEnumerable<CommentNode> comments)
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

        public IEnumerable<IAnnotation> GetModuleAnnotations(IVBComponent component)
        {
            ModuleState result;
            if (_moduleStates.TryGetValue(new QualifiedModuleName(component), out result))
            {
                return result.Annotations;
            }

            return new List<IAnnotation>();
        }

        public void SetModuleAnnotations(IVBComponent component, IEnumerable<IAnnotation> annotations)
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

        internal IDictionary<Tuple<string, DeclarationType>, Attributes> GetModuleAttributes(IVBComponent vbComponent)
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
                if (!declaration.IsBuiltIn)
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

        public bool ClearStateCache(QualifiedModuleName component, bool notifyStateChanged = false)
        {
            var keys = new List<QualifiedModuleName> { component };
            foreach (var key in _moduleStates.Keys)
            {
                if (key.Equals(component) && !keys.Contains(key))
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

        public void AddTokenStream(IVBComponent component, ITokenStream stream)
        {
            _moduleStates[new QualifiedModuleName(component)].SetTokenStream(stream);
        }

        public void AddParseTree(IVBComponent component, IParseTree parseTree)
        {
            var key = new QualifiedModuleName(component);
            _moduleStates[key].SetParseTree(parseTree);
            _moduleStates[key].SetModuleContentHashCode(key.ContentHashCode);
        }

        public IParseTree GetParseTree(IVBComponent component)
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

        public TokenStreamRewriter GetRewriter(IVBComponent component)
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

        public Declaration FindSelectedDeclaration(ICodePane activeCodePane, bool procedureLevelOnly = false)
        {
            if (activeCodePane.IsWrappingNullReference)
            {
                return null;
            }

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
                            item.Item2.ContainsFirstCharacter(selection.Value.Selection) &&
                            item.Item1.DeclarationType != DeclarationType.ModuleOption)
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
            
            return _selectedDeclaration;
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

        public void AddModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule)
        {
            ModuleState referencedModuleState;
            ModuleState referencingModuleState;
            if (!_moduleStates.TryGetValue(referencedModule, out referencedModuleState) || !_moduleStates.TryGetValue(referencingModule, out referencingModuleState))
            {
                return;
            }
            if (referencedModuleState.IsReferencedByModule.Contains(referencingModule))
            {
                return;
            }
            referencedModuleState.IsReferencedByModule.Add(referencingModule);
            referencingModuleState.HasReferenceToModule.AddOrUpdate(referencedModule, 1, (key, value) => value);
        }

        public void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule)
        {
            ModuleState referencingModuleState;
            if (!_moduleStates.TryGetValue(referencingModule, out referencingModuleState))
            {
                return;
            }

            ModuleState referencedModuleState;
            foreach (var referencedModule in referencingModuleState.HasReferenceToModule.Keys)
            {
                if (!_moduleStates.TryGetValue(referencedModule,out referencedModuleState))
                {
                    continue;
                }
                referencedModuleState.IsReferencedByModule.Remove(referencingModule);
            }
            referencingModuleState.RefreshHasReferenceToModule();
        }

        public HashSet<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule)
        { 
            ModuleState referencingModuleState;
            if (!_moduleStates.TryGetValue(referencingModule, out referencingModuleState))
            {
                return new HashSet<QualifiedModuleName>();
            }
            return new HashSet<QualifiedModuleName>(referencingModuleState.HasReferenceToModule.Keys);
        }

        public HashSet<QualifiedModuleName> ModulesReferencedBy(IEnumerable<QualifiedModuleName> referencingModules)
        {
            var referencedModules = new HashSet<QualifiedModuleName>();
            foreach (var referencingModule in referencedModules)
            {
                referencedModules.UnionWith(ModulesReferencedBy(referencingModule));
            }
            return referencedModules;
        }

        public HashSet<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule)
        {
            ModuleState referencedModuleState;
            if (!_moduleStates.TryGetValue(referencedModule, out referencedModuleState))
            {
                return new HashSet<QualifiedModuleName>();
            }
            return new HashSet<QualifiedModuleName>(referencedModuleState.IsReferencedByModule);
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
            _declarationSelections.Clear();
            // no lock because nobody should try to update anything here
            _projects.Clear();

            _isDisposed = true;
        }
    }
}