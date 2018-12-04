using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using NLog;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

// ReSharper disable LoopCanBeConvertedToQuery

namespace Rubberduck.Parsing.VBA
{
    public class ParseProgressEventArgs : EventArgs
    {
        public QualifiedModuleName Module { get; }
        public ParserState State { get; }
        public ParserState OldState { get; }
        public CancellationToken Token { get; }

        public ParseProgressEventArgs(QualifiedModuleName module, ParserState state, ParserState oldState, CancellationToken token)
        {
            Module = module;
            State = state;
            OldState = oldState;
            Token = token;
        }
    }

    public class ParserStateEventArgs : EventArgs
    {
        public ParserStateEventArgs(ParserState state, ParserState oldState, CancellationToken token)
        {
            State = state;
            OldState = oldState;
            Token = token;
        }

        public ParserState State { get; }
        public ParserState OldState { get; }
        public CancellationToken Token { get; }

        public bool IsError => (State == ParserState.ResolverError ||
                                State == ParserState.Error ||
                                State == ParserState.UnexpectedError);
    }

    public enum SuspensionResult
    {
        /// <summary>
        /// The busy action has been queued but has not run yet.
        /// </summary>
        Pending,
        /// <summary>
        /// The busy action was completed successfully.
        /// </summary>
        Completed,
        /// <summary>
        /// The busy action could not executed because it timed out when
        /// attempting to obtain a suspension lock. The timeout is 
        /// governed by the MillisecondsTimeout argument.
        /// </summary>
        TimedOut,
        /// <summary>
        /// The parser arrived to one of states that wasn't listed in the 
        /// AllowedRunStates specified by the requestor (e.g. an error state)
        /// and thus the busy action was not executed.
        /// </summary>
        IncompatibleState,
        /// <summary>
        /// Indicates that the suspension request cannot be made because there 
        /// is no handler for it. This points to a bug in the code.
        /// </summary>
        NotEnabled,
        /// <summary>
        /// An unexpected error; usually indicates a bug in code.
        /// </summary>
        UnexpectedError
    }

    public class RubberduckStatusSuspendParserEventArgs : EventArgs
    {
        public RubberduckStatusSuspendParserEventArgs(object requestor, IEnumerable<ParserState> allowedRunStates, Action busyAction, int millisecondsTimeout)
        {
            Requestor = requestor;
            AllowedRunStates = allowedRunStates;
            BusyAction = busyAction;
            MillisecondsTimeout = millisecondsTimeout;
            Result = SuspensionResult.Pending;
        }

        public object Requestor { get; }
        public IEnumerable<ParserState> AllowedRunStates { get; }
        public Action BusyAction { get; }
        public int MillisecondsTimeout { get; }
        public SuspensionResult Result { get; set; }
    }

    public class RubberduckStatusMessageEventArgs : EventArgs
    {
        public RubberduckStatusMessageEventArgs(string message)
        {
            Message = message;
        }

        public string Message { get; }
    }

    public sealed class RubberduckParserState : IDisposable, IDeclarationFinderProvider, IParseTreeProvider, IParseManager
    {
        public const int NoTimeout = -1;

        private readonly ConcurrentDictionary<QualifiedModuleName, ModuleState> _moduleStates =
            new ConcurrentDictionary<QualifiedModuleName, ModuleState>();

        public event EventHandler<EventArgs> ParseRequest;
        public event EventHandler<RubberduckStatusSuspendParserEventArgs> SuspendRequest;
        public event EventHandler<RubberduckStatusMessageEventArgs> StatusMessageUpdate;

        private static readonly List<ParserState> States = new List<ParserState>();

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public bool IsEnabled { get; internal set; }

        public DeclarationFinder DeclarationFinder { get; private set; }

        private readonly IVBE _vbe;
        private readonly IProjectsRepository _projectRepository;
        private readonly IVBEEvents _vbeEvents;
        private readonly IHostApplication _hostApp;
        private readonly IDeclarationFinderFactory _declarationFinderFactory;

        /// <param name="vbeEvents">Provides event handling from the VBE. Static method <see cref="VBEEvents.Initialize"/> must be already called prior to constructing the method.</param>
        [SuppressMessage("ReSharper", "JoinNullCheckWithUsage")]
        public RubberduckParserState(IVBE vbe, IProjectsRepository projectRepository, IDeclarationFinderFactory declarationFinderFactory, IVBEEvents vbeEvents)
        {
            if (vbe == null)
            {
                throw new ArgumentNullException(nameof(vbe));
            }

            if (projectRepository == null)
            {
                throw new ArgumentException(nameof(projectRepository));
            }

            if (declarationFinderFactory == null)
            {
                throw new ArgumentNullException(nameof(declarationFinderFactory)); 
            }

            if (vbeEvents == null)
            {
                throw new ArgumentNullException(nameof(vbeEvents));
            }

            _vbe = vbe;
            _projectRepository = projectRepository;
            _declarationFinderFactory = declarationFinderFactory;
            _vbeEvents = vbeEvents;
            
            var values = Enum.GetValues(typeof(ParserState));
            foreach (var value in values)
            {
                States.Add((ParserState)value);
            }

            _hostApp = _vbe.HostApplication();
            AddEventHandlers();
            IsEnabled = true;
            RefreshFinder(_hostApp);
        }

        private void RefreshFinder(IHostApplication host)
        {
            var oldDecalarationFinder = DeclarationFinder;
            DeclarationFinder = _declarationFinderFactory.Create(AllDeclarationsFromModuleStates, AllAnnotations, AllUnresolvedMemberDeclarationsFromModulestates, host);
            _declarationFinderFactory.Release(oldDecalarationFinder);
        }

        public void RefreshDeclarationFinder()
        {
            RefreshFinder(_hostApp);
        }

        #region Event Handling

        private void AddEventHandlers()
        {
            _vbeEvents.ProjectAdded += Sinks_ProjectAdded;
            _vbeEvents.ProjectRemoved += Sinks_ProjectRemoved;
            _vbeEvents.ProjectRenamed += Sinks_ProjectRenamed;
            _vbeEvents.ComponentAdded += Sinks_ComponentAdded;
            _vbeEvents.ComponentRemoved += Sinks_ComponentRemoved;
            _vbeEvents.ComponentRenamed += Sinks_ComponentRenamed;
        }

        private void RemoveEventHandlers()
        {
            _vbeEvents.ProjectAdded -= Sinks_ProjectAdded;
            _vbeEvents.ProjectRemoved -= Sinks_ProjectRemoved;
            _vbeEvents.ProjectRenamed -= Sinks_ProjectRenamed;
            _vbeEvents.ComponentAdded -= Sinks_ComponentAdded;
            _vbeEvents.ComponentRemoved -= Sinks_ComponentRemoved;
            _vbeEvents.ComponentRenamed -= Sinks_ComponentRenamed;
        }

        private void Sinks_ProjectAdded(object sender, ProjectEventArgs e)
        {
            if (!_vbe.IsInDesignMode)
            {
                return;
            }

            Logger.Debug("Project '{0}' was added.", e.ProjectId);
            OnParseRequested(sender);
        }

        private void Sinks_ProjectRemoved(object sender, ProjectEventArgs e)
        {
            if (!_vbe.IsInDesignMode)
            {
                return;
            }
            
            Debug.Assert(e.ProjectId != null);
            DisposeProjectDeclarations(e.ProjectId);

            OnParseRequested(sender);
        }

        private void DisposeProjectDeclarations(string projectId)
        {
            var projectDeclarations = DeclarationFinder.UserDeclarations(DeclarationType.Project)
                .Where(declaration => declaration.ProjectId == projectId)
                .OfType<ProjectDeclaration>();
            foreach (var projectDeclaration in projectDeclarations)
            {
                projectDeclaration.Dispose();
            }
        }

        private void Sinks_ProjectRenamed(object sender, ProjectRenamedEventArgs e)
        {
            if (!_vbe.IsInDesignMode || !ThereAreDeclarations())
            {
                return;
            }

            Logger.Debug("Project {0} was renamed.", e.ProjectId);

            OnParseRequested(sender);
        }

        private void Sinks_ComponentAdded(object sender, ComponentEventArgs e)
        {
            if (!_vbe.IsInDesignMode || !ThereAreDeclarations())
            {
                return;
            }

            Logger.Debug("Component '{0}' was added.", e.QualifiedModuleName.ComponentName);
            OnParseRequested(sender);
        }

        private void Sinks_ComponentRemoved(object sender, ComponentEventArgs e)
        {
            if (!_vbe.IsInDesignMode || !ThereAreDeclarations())
            {
                return;
            }

            Logger.Debug("Component '{0}' was removed.", e.QualifiedModuleName.ComponentName);
            OnParseRequested(sender);
        }

        private void Sinks_ComponentRenamed(object sender, ComponentRenamedEventArgs e)
        {
            if (!_vbe.IsInDesignMode || !ThereAreDeclarations())
            {
                return;
            }

            Logger.Debug("Component '{0}' was renamed to '{1}'.", e.OldName, e.QualifiedModuleName.ComponentName);

            //todo: Find out for which situation this drastic (and problematic) cache invalidation has been introduced.
            if (ComponentIsWorksheet(e))
            {
                RefreshProject(e.ProjectId);
                Logger.Debug("Project '{0}' was removed.", e.QualifiedModuleName.ComponentName);
            }
            OnParseRequested(sender);
        }

        private bool ComponentIsWorksheet(ComponentRenamedEventArgs e)
        {
            var componentIsWorksheet = false;
            foreach (var declaration in AllUserDeclarations)
            {
                if (declaration.ProjectId == e.ProjectId &&
                    declaration.DeclarationType == DeclarationType.ClassModule &&
                    declaration.IdentifierName == e.OldName)
                {
                    foreach (var superType in ((ClassModuleDeclaration)declaration).Supertypes)
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

            return componentIsWorksheet;
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
        public void RefreshProjects()
        {
            _projectRepository.Refresh();
        }

        private void RefreshProject(string projectId)
        {
            _projectRepository.Refresh(projectId);

            ClearStateCache(projectId);
        }

        public List<IVBProject> Projects => _projectRepository.Projects().Select(tpl => tpl.Project).ToList();

        public IProjectsProvider ProjectsProvider => _projectRepository;

        public IReadOnlyList<Tuple<QualifiedModuleName, SyntaxErrorException>> ModuleExceptions
        {
            get
            {
                var exceptions = new List<Tuple<QualifiedModuleName, SyntaxErrorException>>();
                foreach (var kvp in _moduleStates)
                {
                    if (kvp.Value.ModuleException == null)
                    {
                        continue;
                    }

                    exceptions.Add(Tuple.Create(kvp.Key, kvp.Value.ModuleException));
                }

                return exceptions;
            }
        }

        public event EventHandler<ParserStateEventArgs> StateChanged;

        private int _stateChangedInvocations;
        private void OnStateChanged(object requestor, CancellationToken token, ParserState state, ParserState oldStatus)
        {
            Interlocked.Increment(ref _stateChangedInvocations);

            Logger.Info($"{nameof(RubberduckParserState)} ({_stateChangedInvocations}) is invoking {nameof(StateChanged)} ({Status})");
            var handler = StateChanged;
            if (handler != null && !token.IsCancellationRequested)
            {
                var args = new ParserStateEventArgs(state, oldStatus, token);
                handler.Invoke(requestor, args);
            }
        }

        public event EventHandler<ParseProgressEventArgs> ModuleStateChanged;

        //Never spawn new threads changing module states in the handler! This will cause deadlocks. 
        private void OnModuleStateChanged(QualifiedModuleName module, ParserState state, ParserState oldState, CancellationToken token)
        {
            var handler = ModuleStateChanged;
            if (handler != null && !token.IsCancellationRequested)
            {
                var args = new ParseProgressEventArgs(module, state, oldState, token);
                handler.Invoke(this, args);
            }
        }

        public void SetModuleState(QualifiedModuleName module, ParserState state, CancellationToken token, SyntaxErrorException parserError = null, bool evaluateOverallState = true)
        {
            if (token.IsCancellationRequested)
            {
                return;
            }

            if (AllUserDeclarations.Any())
            {
                var projectId = module.ProjectId;
                var project = GetProject(projectId);

                if (project == null)
                {
                    // ghost component shouldn't even exist
                    ClearStateCache(module);
                    EvaluateParserState(token);
                    return;
                }
            }

            var oldState = GetModuleState(module);

            _moduleStates.AddOrUpdate(module, new ModuleState(state), (c, e) => e.SetState(state));
            _moduleStates.AddOrUpdate(module, new ModuleState(parserError), (c, e) => e.SetModuleException(parserError));
            Logger.Debug("Module '{0}' state is changing to '{1}' (thread {2})", module.ComponentName, state, Thread.CurrentThread.ManagedThreadId);
            OnModuleStateChanged(module, state, oldState, token);
            if (evaluateOverallState)
            {
                EvaluateParserState(token);
            }
        }

        private IVBProject GetProject(string projectId)
        {
            return _projectRepository.Project(projectId);
        }

        public void EvaluateParserState(CancellationToken token)
        {
            lock (_statusLockObject)
            {
                var newState = OverallParserStateFromModuleStates();
                SetStatusWithCancellation(newState, token);
            }
        }

        private ParserState OverallParserStateFromModuleStates()
        {
            if (_moduleStates.IsEmpty)
            {
                return ParserState.Ready;
            }

            var moduleStates = new List<ParserState>();
            foreach (var moduleState in _moduleStates)
            {
                if (string.IsNullOrEmpty(moduleState.Key.ComponentName) || ProjectsProvider.Component(moduleState.Key) == null || !moduleState.Key.IsParsable)
                {
                    continue;
                }

                moduleStates.Add(moduleState.Value.State);
            }

            if (moduleStates.Count == 0)
            {
                return ParserState.Ready;
            }

            var state = moduleStates[0];
            foreach (var moduleState in moduleStates)
            {
                if (moduleState != moduleStates[0])
                {
                    state = default;
                    break;
                }
            }

            if (state != default)
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
            if (stateCounts[(int)ParserState.UnexpectedError] > 0)
            {
                return ParserState.UnexpectedError;
            }
            if (stateCounts[(int)ParserState.Error] > 0)
            {
                return ParserState.Error;
            }
            if (stateCounts[(int)ParserState.ResolverError] > 0)
            {
                return ParserState.ResolverError;
            }

            //The lowest state wins.
            var result = ParserState.None;
            foreach (var item in moduleStates)
            {
                if (item < result)
                {
                    result = item;
                }
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
        public ParserState Status => _status;

        private void SetStatusWithCancellation(ParserState value, CancellationToken token)
        {
            if (_status != value)
            {
                var oldStatus = _status;
                _status = value;
                OnStateChanged(this, token, _status, oldStatus);
            }
        }

        public void SetStatusAndFireStateChanged(object requestor, ParserState status, CancellationToken token)
        {
            if (Status == status)
            {
                OnStateChanged(requestor, token, status, _status);
            }
            else
            {
                SetStatusWithCancellation(status, token);
            }
        }

        internal void AddModuleStateIfNotPresent(QualifiedModuleName module)
        {
            _moduleStates.AddOrUpdate(module, new ModuleState(ParserState.Pending), (c, s) => s);
        }

        internal void SetModuleAttributes(QualifiedModuleName module, IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> attributes)
        {
            _moduleStates[module].SetModuleAttributes(attributes);
        }

        internal void SetMembersAllowingAttributes(QualifiedModuleName module, IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> membersAllowingAttributes)
        {
            _moduleStates[module].SetMembersAllowingAttributes(membersAllowingAttributes);
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

        public IReadOnlyCollection<CommentNode> GetModuleComments(QualifiedModuleName module)
        {
            if (!_moduleStates.TryGetValue(module, out var moduleState))
            {
                return new List<CommentNode>();
            }

            return moduleState.Comments;
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
            if (_moduleStates.TryGetValue(module, out var result))
            {
                return result.Annotations;
            }

            return Enumerable.Empty<IAnnotation>();
        }

        public void SetModuleAnnotations(QualifiedModuleName module, IEnumerable<IAnnotation> annotations)
        {
            _moduleStates[module].SetAnnotations(new List<IAnnotation>(annotations));
        }

        /// <summary>
        /// Gets a copy of the collected declarations, including the built-in ones.
        /// </summary>
        public IEnumerable<Declaration> AllDeclarations => DeclarationFinder.AllDeclarations;

        /// <summary>
        /// Gets a copy of the collected declarations directly from the module states, including the built-in ones. (Used for refreshing the DeclarationFinder.)
        /// </summary>
        private IReadOnlyList<Declaration> AllDeclarationsFromModuleStates
        {
            get
            {
                var declarations = new List<Declaration>();
                foreach (var state in _moduleStates.Values.Where(state => state.Declarations != null))
                {
                    declarations.AddRange(state.Declarations.Keys);
                }

                return declarations;
            }
        }

        private bool ThereAreDeclarations()
        {
            return _moduleStates.Values.Any(state => state.Declarations != null && state.Declarations.Keys.Any());
        }

        /// <summary>
        /// Gets a copy of the unresolved member declarations directly from the module states. (Used for refreshing the DeclarationFinder.)
        /// </summary>
        private IReadOnlyList<UnboundMemberDeclaration> AllUnresolvedMemberDeclarationsFromModulestates
        {
            get
            {
                var declarations = new List<UnboundMemberDeclaration>();
                foreach (var state in _moduleStates.Values.Where(state => state.UnresolvedMemberDeclarations != null))
                {
                    declarations.AddRange(state.UnresolvedMemberDeclarations.Keys);
                }

                return declarations;
            }
        }

        /// <summary>
        /// Gets a copy of the collected declarations, excluding the built-in ones.
        /// </summary>
        public IEnumerable<Declaration> AllUserDeclarations => DeclarationFinder.AllUserDeclarations;

        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> GetModuleAttributes(QualifiedModuleName module)
        {
            return _moduleStates[module].ModuleAttributes;
        }

        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> GetMembersAllowingAttributes(QualifiedModuleName module)
        {
            return _moduleStates[module].MembersAllowingAttributes;
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
                while (!declarations.TryRemove(declaration, out var _))
                {
                    Logger.Warn("Could not remove existing unresolved member declaration for '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
                }
            }
            while (!declarations.TryAdd(declaration, 0) && !declarations.ContainsKey(declaration))
            {
                Logger.Warn("Could not add unresolved member declaration '{0}' ({1}). Retrying.", declaration.IdentifierName, declaration.DeclarationType);
            }
        }

        public void ClearStateCache(string projectId)
        {
            try
            {
                foreach (var moduleState in _moduleStates.Where(moduleState => moduleState.Key.ProjectId == projectId))
                {
                    var qualifiedModuleName = moduleState.Key;
                    if (qualifiedModuleName.ComponentType == ComponentType.Undefined || qualifiedModuleName.ComponentType == ComponentType.ComComponent)
                    {
                        if (_moduleStates.TryRemove(qualifiedModuleName, out var state))
                        {
                            state.Dispose();
                        }
                    }
                    else
                    {
                        //This should be a user component.
                        while (!ClearStateCache(qualifiedModuleName))
                        {
                            // until Hell freezes over?
                        }
                    }
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception, $"Unexpected COMException while clearing the project with projectId {projectId}. Clearing all modules.");
                _moduleStates.Clear();
            }
        }


        public bool ClearStateCache(QualifiedModuleName module)
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

            return success;
        }

        private bool RemoveKeysFromCollections(IEnumerable<QualifiedModuleName> keys)
        {
            var success = true;
            foreach (var key in keys)
            {
                ModuleState moduleState = null;
                success = success && (!_moduleStates.ContainsKey(key) || _moduleStates.TryRemove(key, out moduleState));
                moduleState?.Dispose();
            }

            return success;
        }

        public void SetCodePaneTokenStream(QualifiedModuleName module, ITokenStream codePaneTokenStream)
        {
            _moduleStates[module].SetCodePaneTokenStream(codePaneTokenStream);
        }

        public void SaveContentHash(QualifiedModuleName module)
        {
            _moduleStates[module].SetModuleContentHashCode(GetModuleContentHash(module));
        }

        public void AddParseTree(QualifiedModuleName module, IParseTree parseTree, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var key = module;
            _moduleStates[key].SetParseTree(parseTree, codeKind);
        }

        public IParseTree GetParseTree(QualifiedModuleName module, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            switch (codeKind)
            {
                case CodeKind.AttributesCode:
                    return _moduleStates[module].AttributesPassParseTree;
                case CodeKind.CodePaneCode:
                    return _moduleStates[module].ParseTree;
                default:
                    throw new ArgumentOutOfRangeException(nameof(codeKind), codeKind, null);
            }
        }

        public List<KeyValuePair<QualifiedModuleName, IParseTree>> AttributeParseTrees
        {
            get
            {
                var parseTrees = new List<KeyValuePair<QualifiedModuleName, IParseTree>>();
                foreach(var state in _moduleStates)
                {
                    if(state.Value.AttributesPassParseTree != null)
                    {
                        parseTrees.Add(new KeyValuePair<QualifiedModuleName, IParseTree>(state.Key, state.Value.AttributesPassParseTree));
                    }
                }

                return parseTrees;
            }
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
                        parseTrees.Add(new KeyValuePair<QualifiedModuleName, IParseTree>(state.Key, state.Value.ParseTree));
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
                    using (var components = project.VBComponents)
                    {
                        foreach (var component in components)
                        using (component)
                        {
                            if (IsNewOrModified(component))
                            {
                                return true;
                            }
                        }
                    }
                }
                catch (COMException)
                {
                }
            }

            return false;
        }

        public ITokenStream GetCodePaneTokenStream(QualifiedModuleName qualifiedModuleName)
        {
            return _moduleStates[qualifiedModuleName].CodePaneTokenStream;
        }

        public ITokenStream GetAttributesTokenStream(QualifiedModuleName qualifiedModuleName)
        {
            return _moduleStates[qualifiedModuleName].AttributesTokenStream;
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
        public void OnParseRequested(object requestor)
        {
            var handler = ParseRequest;
            if (handler != null && IsEnabled)
            {
                var args = EventArgs.Empty;
                handler.Invoke(requestor, args);
            }
        }

        public SuspensionResult OnSuspendParser(object requestor, IEnumerable<ParserState> allowedRunStates, Action busyAction, int millisecondsTimeout = NoTimeout)
        {
            if (millisecondsTimeout < NoTimeout)
            {
                throw new ArgumentOutOfRangeException(nameof(millisecondsTimeout));
            }

            var handler = SuspendRequest;
            if (handler != null && IsEnabled)
            {
                var args = new RubberduckStatusSuspendParserEventArgs(requestor, allowedRunStates, busyAction, millisecondsTimeout);
                handler.Invoke(requestor, args);
                return args.Result;
            }

            return SuspensionResult.NotEnabled;
        }

        public bool IsNewOrModified(IVBComponent component)
        {
            var key = new QualifiedModuleName(component);
            return IsNewOrModified(key);
        }

        public bool IsNewOrModified(QualifiedModuleName key)
        {
            if (_moduleStates.TryGetValue(key, out var moduleState))
            {
                // existing/modified
                return moduleState.IsNew || GetModuleContentHash(key) != moduleState.ModuleContentHashCode;
            }

            // new
            return true;
        }

        private int GetModuleContentHash(QualifiedModuleName module)
        {
            var component = ProjectsProvider.Component(module);
            return QualifiedModuleName.GetContentHash(component);
        }

        public Declaration FindSelectedDeclaration(ICodePane activeCodePane)
        {
            if (activeCodePane != null)
            {
                return DeclarationFinder?.FindSelectedDeclaration(activeCodePane);
            }

            using (var active = _vbe.ActiveCodePane)
            {
                return DeclarationFinder?.FindSelectedDeclaration(active);
            }
        }

        public void RemoveBuiltInDeclarations(QualifiedModuleName projectQmn)
        {
            ClearAsTypeDeclarationPointingToReference(projectQmn);
            if (_moduleStates.TryRemove(projectQmn, out var moduleState))
            {
                moduleState?.Dispose();
                Logger.Warn("Could not remove declarations for removed reference '{0}' ({1}).", projectQmn.Name, projectQmn.ProjectId);
            }
        }
        
        private void ClearAsTypeDeclarationPointingToReference(QualifiedModuleName referencedProject)
        {
            var toClearAsTypeDeclaration = DeclarationFinder
                                            .FindDeclarationsWithNonBaseAsType()
                                            .Where(decl => decl.QualifiedName.QualifiedModuleName == referencedProject);
            foreach(var declaration in toClearAsTypeDeclaration)
            {
                declaration.AsTypeDeclaration = null;
            }
        }

        public void SetAttributesTokenStream(QualifiedModuleName module, ITokenStream attributesTokenStream)
        {
            _moduleStates[module].SetAttributesTokenStream(attributesTokenStream);
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

            RemoveEventHandlers();
            VBEEvents.Terminate();

            _moduleStates.Clear();

            // no lock because nobody should try to update anything here
            _projectRepository.Dispose();

            _isDisposed = true;
        }
    }
}