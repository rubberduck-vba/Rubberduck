using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.DeclarationCaching
{
    public class DeclarationFinder
    {
        private static readonly SquareBracketedNameComparer NameComparer = new SquareBracketedNameComparer();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IHostApplication _hostApp;
        private readonly AnnotationService _annotationService;
        private IDictionary<string, List<Declaration>> _declarationsByName;
        private IDictionary<QualifiedModuleName, List<Declaration>> _declarations;
        private readonly ConcurrentDictionary<QualifiedMemberName, ConcurrentBag<Declaration>> _newUndeclared;
        private readonly ConcurrentBag<UnboundMemberDeclaration> _newUnresolved;
        private List<UnboundMemberDeclaration> _unresolved;
        private IDictionary<QualifiedModuleName, List<IAnnotation>> _annotations;
        private IDictionary<Declaration, List<ParameterDeclaration>> _parametersByParent;
        private IDictionary<DeclarationType, List<Declaration>> _userDeclarationsByType;
        private IDictionary<QualifiedSelection, List<Declaration>> _declarationsBySelection;
        private IDictionary<QualifiedSelection, List<IdentifierReference>> _referencesBySelection;
        private IReadOnlyDictionary<QualifiedModuleName, IReadOnlyList<IdentifierReference>> _referencesByModule;
        private IDictionary<QualifiedMemberName, List<IdentifierReference>> _referencesByMember;

        private Lazy<IDictionary<DeclarationType, List<Declaration>>> _builtInDeclarationsByType;
        private Lazy<IDictionary<Declaration, List<Declaration>>> _handlersByWithEventsField;

        private Lazy<IDictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>>> _implementingMembers;
        private Lazy<IDictionary<VBAParser.ImplementsStmtContext, List<ModuleBodyElementDeclaration>>> _membersByImplementsContext;
        private Lazy<IDictionary<ClassModuleDeclaration, List<Declaration>>> _interfaceMembers;
        private Lazy<IDictionary<ClassModuleDeclaration, List<ClassModuleDeclaration>>> _interfaceImplementions;
        private Lazy<IDictionary<IInterfaceExposable, List<ModuleBodyElementDeclaration>>> _implementationsByMember;

        private Lazy<List<Declaration>> _nonBaseAsType;
        private Lazy<List<Declaration>> _eventHandlers;
        private Lazy<List<Declaration>> _projects;
        private Lazy<List<Declaration>> _classes;
        
        private static QualifiedSelection GetGroupingKey(Declaration declaration)
        {
            // we want the procedures' whole body, not just their identifier:
            return declaration.DeclarationType.HasFlag(DeclarationType.Member)
                ? new QualifiedSelection(
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.Context.GetSelection())
                : declaration.QualifiedSelection;
        }

        public DeclarationFinder(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, 
            IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, IHostApplication hostApp = null)
        {
            _hostApp = hostApp;

            _newUndeclared = new ConcurrentDictionary<QualifiedMemberName, ConcurrentBag<Declaration>>(new Dictionary<QualifiedMemberName, ConcurrentBag<Declaration>>());
            _newUnresolved = new ConcurrentBag<UnboundMemberDeclaration>(new List<UnboundMemberDeclaration>());

            _annotationService = new AnnotationService(this);

            var collectionConstructionActions = CollectionConstructionActions(declarations, annotations, unresolvedMemberDeclarations);
            ExecuteCollectionConstructionActions(collectionConstructionActions);

            //Temporal coupling: the initializers of the lazy collections use the regular collections filled above.
            InitializeLazyCollections();
        }

        protected virtual void ExecuteCollectionConstructionActions(List<Action> collectionConstructionActions)
        {
            collectionConstructionActions.ForEach(action => action.Invoke());
        }

        private List<Action> CollectionConstructionActions(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, 
            IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations)
        {
            var actions = new List<Action>
            {
                () =>
                    _unresolved = unresolvedMemberDeclarations
                        .ToList(),
                () =>
                    _annotations = annotations
                        .GroupBy(node => node.QualifiedSelection.QualifiedName)
                        .ToDictionary(),
                () =>
                    _declarations = declarations
                        .GroupBy(item => item.QualifiedName.QualifiedModuleName)
                        .ToDictionary(),
                () =>
                    _declarationsByName = declarations
                        .GroupBy(declaration => declaration.IdentifierName.ToLowerInvariant())
                        .ToDictionary(),
                () =>
                    _declarationsBySelection = declarations
                        .Where(declaration => declaration.IsUserDefined)
                        .GroupBy(GetGroupingKey)
                        .ToDictionary(),
                () =>
                    _referencesBySelection = declarations
                        .SelectMany(declaration => declaration.References)
                        .GroupBy(
                            reference => new QualifiedSelection(reference.QualifiedModuleName, reference.Selection))
                        .ToDictionary(),
                () =>
                    _parametersByParent = declarations
                        .Where(declaration => declaration.DeclarationType == DeclarationType.Parameter)
                        .Cast<ParameterDeclaration>()
                        .GroupBy(declaration => declaration.ParentDeclaration)
                        .ToDictionary(),
                () =>
                    _userDeclarationsByType = declarations
                        .Where(declaration => declaration.IsUserDefined)
                        .GroupBy(declaration => declaration.DeclarationType)
                        .ToDictionary(),
                () =>
                    _referencesByModule = declarations
                        .SelectMany(declaration => declaration.References)
                        .GroupBy(reference =>
                            Declaration.GetModuleParent(reference.ParentScoping).QualifiedName.QualifiedModuleName)
                        .ToReadonlyDictionary(),
                () =>
                    _referencesByMember = declarations
                        .SelectMany(declaration => declaration.References)
                        .GroupBy(reference => reference.ParentScoping.QualifiedName)
                        .ToDictionary()
            };

            return actions;
        }

        private void InitializeLazyCollections()
        {
            _builtInDeclarationsByType = new Lazy<IDictionary<DeclarationType, List<Declaration>>>(() =>
                _declarations
                    .AllValues()
                    .Where(declaration => !declaration.IsUserDefined)
                    .GroupBy(declaration => declaration.DeclarationType)
                    .ToDictionary()
                , true);

            _nonBaseAsType = new Lazy<List<Declaration>>(() =>
                _declarations
                    .AllValues()
                    .Where(d => !string.IsNullOrWhiteSpace(d.AsTypeName)
                                    && !d.AsTypeIsBaseType
                                    && d.DeclarationType != DeclarationType.Project
                                    && d.DeclarationType != DeclarationType.ProceduralModule)
                    .ToList()
                , true);

            _eventHandlers = new Lazy<List<Declaration>>(FindAllEventHandlers, true);
            _projects = new Lazy<List<Declaration>>(() => DeclarationsWithType(DeclarationType.Project).ToList(), true);
            _classes = new Lazy<List<Declaration>>(() => DeclarationsWithType(DeclarationType.ClassModule).ToList(), true);
            _handlersByWithEventsField = new Lazy<IDictionary<Declaration, List<Declaration>>>(FindAllHandlersByWithEventField, true);

            _implementingMembers = new Lazy<IDictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>>>(FindAllImplementingMembers, true);
            _interfaceMembers = new Lazy<IDictionary<ClassModuleDeclaration, List<Declaration>>>(FindAllIinterfaceMembersByModule, true);
            _membersByImplementsContext = new Lazy<IDictionary<VBAParser.ImplementsStmtContext, List<ModuleBodyElementDeclaration>>>(FindAllImplementingMembersByImplementsContext, true);
            _interfaceImplementions = new Lazy<IDictionary<ClassModuleDeclaration, List<ClassModuleDeclaration>>>(FindAllImplementionsByInterface, true);
            _implementationsByMember = new Lazy<IDictionary<IInterfaceExposable, List<ModuleBodyElementDeclaration>>>(FindAllImplementingMembersByMember, true);
        }

        private IDictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>> FindAllImplementingMembers()
        {
            var implementsInstructions = UserDeclarations(DeclarationType.ClassModule)
                .SelectMany(cls => cls.References
                    .Where(reference => reference.Context is VBAParser.ImplementsStmtContext 
                        || (reference.Context).IsDescendentOf<VBAParser.ImplementsStmtContext>())
                    .Select(reference =>
                        new
                        {
                            IdentifierReference = reference,
                            Context = reference.Context is VBAParser.ImplementsStmtContext context 
                                ? context 
                                : reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>()
                        }
                    )
                ).ToList();

            var output = new Dictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>>();
            foreach (var impl in implementsInstructions)
            {
                output.Add((impl.Context, impl.IdentifierReference.ParentScoping),
                    ((ClassModuleDeclaration) impl.IdentifierReference.ParentScoping).Members.Where(item =>
                        item is ModuleBodyElementDeclaration member && ReferenceEquals(member.InterfaceImplemented,
                            impl.IdentifierReference.Declaration))
                    .Cast<ModuleBodyElementDeclaration>().ToList());
            }

            return output;
        }

        private Dictionary<ClassModuleDeclaration, List<ClassModuleDeclaration>> FindAllImplementionsByInterface()
        {
            return UserDeclarations(DeclarationType.ClassModule)
                .Cast<ClassModuleDeclaration>()
                .Where(module => module.IsInterface).ToDictionary(intrface => intrface,
                    intrface => intrface.Subtypes.Cast<ClassModuleDeclaration>()
                        .Where(type => type.ImplementedInterfaces.Contains(intrface)).ToList());
        }

        private IDictionary<IInterfaceExposable, List<ModuleBodyElementDeclaration>> FindAllImplementingMembersByMember()
        {
            var implementations = _implementingMembers.Value.AllValues();
            return implementations
                
                .GroupBy(member => (IInterfaceExposable)member.InterfaceMemberImplemented)
                .ToDictionary(member => member.Key, member => member.ToList());
        }

        private IDictionary<VBAParser.ImplementsStmtContext, List<ModuleBodyElementDeclaration>> FindAllImplementingMembersByImplementsContext()
        {
            return _implementingMembers.Value.ToDictionary(pair => pair.Key.Context, pair => pair.Value);
        }

        private IDictionary<ClassModuleDeclaration, List<Declaration>> FindAllIinterfaceMembersByModule()
        {
            return UserDeclarations(DeclarationType.ClassModule)
                .Cast<ClassModuleDeclaration>()
                .Where(module => module.IsInterface)
                .ToDictionary(
                    module => module,
                    module => module.Members
                        .Where(member => member is IInterfaceExposable candidate && candidate.IsInterfaceMember)
                        .ToList());
        }

        private IDictionary<Declaration, List<Declaration>> FindAllHandlersByWithEventField()
        {
            var withEventsFields = UserDeclarations(DeclarationType.Variable).Where(item => item.IsWithEvents);
            var events = withEventsFields.Select(field =>
                new
                {
                    WithEventsField = field,
                    AvailableEvents = FindEvents(field.AsTypeDeclaration).ToArray()
                });

            var handlersByWithEventsField = events.Select(item =>
                    new
                    {
                        item.WithEventsField,
                        Handlers = item.AvailableEvents.SelectMany(evnt =>
                            Members(item.WithEventsField.ParentDeclaration.QualifiedName.QualifiedModuleName)
                                .Where(member => member.DeclarationType == DeclarationType.Procedure
                                                && member.IdentifierName == item.WithEventsField.IdentifierName + "_" + evnt.IdentifierName))
                    })
                    .ToDictionary(item => item.WithEventsField, item => item.Handlers.ToList());
            return handlersByWithEventsField;
        }

        public Declaration FindSelectedDeclaration(ICodePane activeCodePane)
        {
            if (activeCodePane == null || activeCodePane.IsWrappingNullReference)
            {
                return null;
            }
            
            var qualifiedSelection = activeCodePane.GetQualifiedSelection();
            if (!qualifiedSelection.HasValue || qualifiedSelection.Value.Equals(default))
            {
                return null;
            }

            var selection = qualifiedSelection.Value.Selection;

            // statistically we'll be on an IdentifierReference more often than on a Declaration:
            var matches = _referencesBySelection
                .Where(kvp => kvp.Key.QualifiedName.Equals(qualifiedSelection.Value.QualifiedName)
                    && kvp.Key.Selection.ContainsFirstCharacter(qualifiedSelection.Value.Selection))
                .SelectMany(kvp => kvp.Value)
                .OrderByDescending(reference => reference.Declaration.DeclarationType)
                .Select(reference => reference.Declaration)
                .Distinct()
                .ToArray();

            if (!matches.Any())
            {
                matches = _declarationsBySelection
                    .Where(kvp => kvp.Key.QualifiedName.Equals(qualifiedSelection.Value.QualifiedName)
                        && kvp.Key.Selection.ContainsFirstCharacter(selection))
                    .SelectMany(kvp => kvp.Value)
                    .OrderByDescending(declaration => declaration.DeclarationType)
                    .Distinct()
                    .ToArray();
            }

            switch (matches.Length)
            {
                case 0:
                    return ModuleDeclaration(qualifiedSelection.Value.QualifiedName);

                case 1:
                    return matches.Single();

                default:
                    // they're sorted by type, so a local comes before the procedure it's in
                    return matches.FirstOrDefault();
            }
        }

        public IEnumerable<Declaration> FreshUndeclared => _newUndeclared.AllValues();

        //This does not need a lock because enumerators over a ConcurrentBag uses a snapshot.    
        public IEnumerable<UnboundMemberDeclaration> FreshUnresolvedMemberDeclarations => _newUnresolved.ToList();

        public IEnumerable<UnboundMemberDeclaration> UnresolvedMemberDeclarations => _unresolved;

        public IEnumerable<Declaration> Members(Declaration module)
        {
            return Members(module.QualifiedName.QualifiedModuleName);
        }

        public IEnumerable<Declaration> Members(QualifiedModuleName module)
        {
            return _declarations.TryGetValue(module, out var members)
                    ? members
                    : Enumerable.Empty<Declaration>();
        }

        public Declaration ModuleDeclaration(QualifiedModuleName module)
        {
            return Members(module).SingleOrDefault(member => member.DeclarationType.HasFlag(DeclarationType.Module));
        }

        public IReadOnlyCollection<QualifiedModuleName> AllModules => _declarations.Keys.AsReadOnly();

        public IEnumerable<Declaration> AllDeclarations => _declarations.AllValues();

        public IEnumerable<Declaration> FindDeclarationsWithNonBaseAsType()
        {
            return _nonBaseAsType.Value;
        }
 
        public IEnumerable<Declaration> FindEventHandlers()
        {
            return _eventHandlers.Value;
        }

        public IEnumerable<Declaration> Classes => _classes.Value;
        public IEnumerable<Declaration> Projects => _projects.Value;

        public IEnumerable<Declaration> UserDeclarations(DeclarationType type)
        {
            return _userDeclarationsByType.TryGetValue(type, out var result)
                ? result
                : _userDeclarationsByType
                    .Where(item => item.Key.HasFlag(type))
                    .SelectMany(item => item.Value);
        }

        public IEnumerable<Declaration> AllUserDeclarations => _userDeclarationsByType.AllValues();

        public IEnumerable<Declaration> BuiltInDeclarations(DeclarationType type)
        {
            return _builtInDeclarationsByType.Value.TryGetValue(type, out var result)
                ? result
                : _builtInDeclarationsByType.Value
                    .Where(item => item.Key.HasFlag(type))
                    .SelectMany(item => item.Value);
        }

        public IEnumerable<Declaration> AllBuiltInDeclarations => _builtInDeclarationsByType.Value.AllValues();

        public IEnumerable<Declaration> DeclarationsWithType(DeclarationType type)
        {
            return BuiltInDeclarations(type).Concat(UserDeclarations(type));
        }

        public IEnumerable<Declaration> FindHandlersForWithEventsField(Declaration field)
        {
            return _handlersByWithEventsField.Value.TryGetValue(field, out var result) 
                ? result 
                : Enumerable.Empty<Declaration>();
        }

        /// <summary>
        /// Finds all members of a class that are implementing the interface defined by the passed context.
        /// </summary>
        /// <param name="context">The ImplementsStmtContext to find member for, e.g. 'Implements IFoo`</param>
        /// <returns>Members of the containing class that implement the interface.</returns>
        public IEnumerable<Declaration> FindInterfaceMembersForImplementsContext(VBAParser.ImplementsStmtContext context)
        {
            return _membersByImplementsContext.Value.TryGetValue(context, out var result)
                ? result
                : Enumerable.Empty<Declaration>();
        }

        /// <summary>
        /// Finds the interface declaration for a QualifiedSelection contained by a statement similar to "Implements IClass1"
        /// </summary>
        /// <param name="selection">The QualifiedSelection to search.</param>
        /// <returns>The selected interface if found, null if not found.</returns>
        public ClassModuleDeclaration FindInterface(QualifiedSelection selection)
        {
            return FindAllUserInterfaces()
                .FirstOrDefault(declaration => declaration.References
                    .Any(reference => reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>() != null 
                                      && ReferenceEquals(reference.Declaration, declaration)));
        }

        /// <summary>
        /// Finds all user interface definition declarations.
        /// </summary>
        /// <returns>All user interface definition declarations.</returns>
        public IEnumerable<ClassModuleDeclaration> FindAllUserInterfaces()
        {
            return _interfaceMembers.Value.Keys;
        }

        /// <summary>
        /// Finds all classes that implement the passed user interface.
        /// </summary>
        /// <param name="interfaceDeclaration">The interface to find implementations of.</param>
        /// <returns>All classes implementing the interface.</returns>
        public IEnumerable<Declaration> FindAllImplementationsOfInterface(ClassModuleDeclaration interfaceDeclaration)
        {
            var lookup = _interfaceImplementions.Value;
            return lookup.TryGetValue(interfaceDeclaration, out var implementations)
                ? implementations
                : Enumerable.Empty<Declaration>();
        }

        /// <summary>
        /// Finds all members of user interfaces.
        /// </summary>
        /// <returns>All members of user interfaces.</returns>
        public IEnumerable<Declaration> FindAllInterfaceMembers()
        {
            return _interfaceMembers.Value.SelectMany(item => item.Value);
        }

        /// <summary>
        /// Finds all concrete implementations of interface members.
        /// </summary>
        /// <returns>All declarations that implement an interface member.</returns>
        public IEnumerable<ModuleBodyElementDeclaration> FindAllInterfaceImplementingMembers()
        {
            return _membersByImplementsContext.Value.AllValues().Distinct();
        }

        /// <summary>
        /// Locates all concrete implementations of the passed interface declaration.
        /// </summary>
        /// <param name="interfaceDeclaration">The interface member to find.</param>
        /// <returns>All concrete implementations of the passed interface declaration.</returns>
        public IEnumerable<ModuleBodyElementDeclaration> FindInterfaceImplementationMembers(Declaration interfaceMember)
        {
            if (!(interfaceMember is IInterfaceExposable member))
            {
                return Enumerable.Empty<ModuleBodyElementDeclaration>();
            }

            return _implementationsByMember.Value.TryGetValue(member, out var implementations)
                ? implementations
                : Enumerable.Empty<ModuleBodyElementDeclaration>();
        }

        public ParameterDeclaration FindParameter(Declaration procedure, string parameterName)
        {
            return _parametersByParent.TryGetValue(procedure, out List<ParameterDeclaration> parameters) 
                ? parameters.SingleOrDefault(parameter => parameter.IdentifierName == parameterName) 
                : null;
        }

        public IEnumerable<ParameterDeclaration> Parameters(Declaration procedure)
        {
            return _parametersByParent.TryGetValue(procedure, out List<ParameterDeclaration> result)
                ? result
                : Enumerable.Empty<ParameterDeclaration>();
        }

        public IEnumerable<Declaration> FindMemberMatches(Declaration parent, string memberName)
        {
            return _declarations.TryGetValue(parent.QualifiedName.QualifiedModuleName, out var children)
                ? children.Where(item => item.DeclarationType.HasFlag(DeclarationType.Member)
                                             && item.IdentifierName == memberName)
                : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<IAnnotation> FindAnnotations(QualifiedModuleName module)
        {
            return _annotations.TryGetValue(module, out var result) 
                ? result 
                : Enumerable.Empty<IAnnotation>();
        }

        public bool IsMatch(string declarationName, string potentialMatchName)
        {
            return string.Equals(declarationName, potentialMatchName, StringComparison.OrdinalIgnoreCase);
        }

        private IEnumerable<Declaration> FindEvents(Declaration module)
        {
            Debug.Assert(module != null);

            var members = Members(module.QualifiedName.QualifiedModuleName);
            return members == null 
                ? Enumerable.Empty<Declaration>() 
                : members.Where(declaration => declaration.DeclarationType == DeclarationType.Event);
        }

        public Declaration FindEvent(Declaration module, string eventName)
        {
            var matches = MatchName(eventName);
            return matches.FirstOrDefault(m => module.Equals(Declaration.GetModuleParent(m)) && m.DeclarationType == DeclarationType.Event);
        }

        public Declaration FindLabel(Declaration procedure, string label)
        {
            var matches = MatchName(label);
            return matches.FirstOrDefault(m => procedure.Equals(m.ParentDeclaration) && m.DeclarationType == DeclarationType.LineLabel);
        }

        public IEnumerable<Declaration> MatchName(string name)
        {
            var normalizedName = ToNormalizedName(name);
            return _declarationsByName.TryGetValue(normalizedName, out var result) 
                ? result 
                : Enumerable.Empty<Declaration>();
        }

        public ParameterDeclaration FindParameterFromArgument(VBAParser.ArgumentExpressionContext argExpression, Declaration enclosingProcedure)
        {
            if (argExpression  == null || 
                argExpression.GetDescendent<VBAParser.ParenthesizedExprContext>() != null || 
                argExpression.BYVAL() != null)
            {
                // not an argument, or argument is parenthesized and thus passed ByVal
                return null;
            }

            var callStmt = argExpression?.GetAncestor<VBAParser.CallStmtContext>();
            var procedureName = callStmt?.GetDescendent<VBAParser.LExpressionContext>()
                                         .GetDescendents<VBAParser.IdentifierContext>()
                                         .LastOrDefault()?.GetText();
            if (procedureName == null)
            {
                // if we don't know what we're calling, we can't dig any further
                return null;
            }

            var procedure = MatchName(procedureName)
                .Where(p => AccessibilityCheck.IsAccessible(enclosingProcedure, p))
                .SingleOrDefault(p => !p.DeclarationType.HasFlag(DeclarationType.Property) || p.DeclarationType.HasFlag(DeclarationType.PropertyGet));
            if (procedure?.ParentScopeDeclaration is ClassModuleDeclaration)
            {
                // we can't know that the member is on the class' default interface
                return null;
            }

            var parameters = Parameters(procedure);

            ParameterDeclaration parameter;
            var namedArg = argExpression.GetAncestor<VBAParser.NamedArgumentContext>();
            if (namedArg != null)
            {
                // argument is named: we're lucky
                var parameterName = namedArg.unrestrictedIdentifier().GetText();
                parameter = parameters.SingleOrDefault(p => p.IdentifierName == parameterName);
            }
            else
            {
                // argument is positional: work out its index
                var argList = callStmt.GetDescendent<VBAParser.ArgumentListContext>();
                var args = argList.GetDescendents<VBAParser.PositionalArgumentContext>().ToArray();

                var parameterIndex = args
                    .Select((param, index) => param.GetDescendent<VBAParser.ArgumentExpressionContext>() == argExpression ? (param, index) : (null, -1))
                    .SingleOrDefault(item => item.param != null).index;

                parameter = parameters
                    .OrderBy(p => p.Selection)
                    .Select((param, index) => (param, index))
                    .SingleOrDefault(item => item.index == parameterIndex).param;
            }

            return parameter;
        }

        private string ToNormalizedName(string name)
        {
            var lower = name.ToLowerInvariant();
            if (lower.Length > 1 && lower[0] == '[' && lower[lower.Length - 1] == ']')
            {
                var result = lower.Substring(1, lower.Length - 2);
                return result;
            }
            return lower;
        }

        public Declaration FindProject(string name, Declaration currentScope = null)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(project => 
                    project.DeclarationType.HasFlag(DeclarationType.Project)
                    && (currentScope == null || project.ProjectId == currentScope.ProjectId));
            }
            catch (InvalidOperationException exception)
            {
                Logger.Error(exception, "Multiple matches found for project '{0}'.", name);
            }

            return result;
        }

        public Declaration FindStdModule(string name, Declaration parent, bool includeBuiltIn = false)
        {
            Debug.Assert(parent != null);
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.ProceduralModule)
                    && (parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || declaration.IsUserDefined));
            }
            catch (InvalidOperationException exception)
            {
                Logger.Error(exception, "Multiple matches found for std.module '{0}'.", name);
            }

            return result;
        }

        public Declaration FindClassModule(string name, Declaration parent, bool includeBuiltIn = false)
        {
            Debug.Assert(parent != null);
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.ClassModule)
                    && (parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || declaration.IsUserDefined));
            }
            catch (InvalidOperationException exception)
            {
                Logger.Error(exception, "Multiple matches found for class module '{0}'.", name);
            }

            return result;
        }

        public Declaration FindReferencedProject(Declaration callingProject, string referencedProjectName)
        {
            return FindInReferencedProjectByPriority(callingProject, referencedProjectName, p => p.DeclarationType.HasFlag(DeclarationType.Project));
        }

        public Declaration FindModuleEnclosingProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, string calleeModuleName, DeclarationType moduleType)
        {
            var nameMatches = MatchName(calleeModuleName);
            var moduleMatches = nameMatches.Where(m =>
                m.DeclarationType.HasFlag(moduleType)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && !m.Equals(callingModule));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassEnclosingProject(Declaration callingProject, Declaration callingModule, string defaultInstanceVariableClassName)
        {
            var nameMatches = MatchName(defaultInstanceVariableClassName);
            var moduleMatches = nameMatches.Where(m =>
                m.DeclarationType == DeclarationType.ClassModule && ((ClassModuleDeclaration)m).HasDefaultInstanceVariable
                && Declaration.GetProjectParent(m).Equals(callingProject)).ToList(); 
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, 
            string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName,
                p => referencedProject.Equals(Declaration.GetProjectParent(p)) &&
                     p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName,
                p => p.DeclarationType == DeclarationType.ClassModule &&
                     ((ClassModuleDeclaration) p).HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, 
            string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName,
                p => referencedProject.Equals(Declaration.GetProjectParent(p))
                    && p.DeclarationType == DeclarationType.ClassModule 
                    && ((ClassModuleDeclaration)p).HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberWithParent(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration parent, 
            string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && parent.Equals(m.ParentDeclaration))
                .ToList();
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            if (match != null)
            {
                return match;
            }
            return ClassModuleDeclaration
                .GetSupertypes(parent)
                .Select(supertype => 
                    FindMemberWithParent(callingProject, callingModule, callingParent, supertype, memberName, memberType))
                .FirstOrDefault(supertypeMember => supertypeMember != null);
        }

        public Declaration FindMemberEnclosingModule(Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            // We do not explicitly pass the callingProject here because we have to walk up the type hierarchy
            // and thus the project differs depending on the callingModule.
            var callingProject = Declaration.GetProjectParent(callingModule);
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && Declaration.GetProjectParent(m).Equals(callingProject)
                            && callingModule.Equals(Declaration.GetModuleParent(m))
                ).ToList();
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            if (match != null)
            {
                return match;
            }
            // Classes such as Worksheet have properties such as Range that can be access in a user defined class such as Sheet1,
            // that's why we have to walk the type hierarchy and find these implementations.
            foreach (var supertype in ClassModuleDeclaration.GetSupertypes(callingModule))
            {
                // Only built-in classes such as Worksheet can be considered "real base classes".
                // User created interfaces work differently and don't allow accessing accessing implementations.
                if (supertype.IsUserDefined)
                {
                    continue;
                }
                var supertypeMatch = FindMemberEnclosingModule(supertype, callingParent, memberName, memberType);
                if (supertypeMatch != null)
                {
                    return supertypeMatch;
                }
            }

            return null;
        }

        public Declaration FindMemberEnclosingProcedure(Declaration enclosingProcedure, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && enclosingProcedure.Equals(m.ParentDeclaration)).ToList();

            if (memberMatches.Any())
            {
                return memberMatches.FirstOrDefault();
            }

            if (memberType == DeclarationType.Variable && NameComparer.Equals(enclosingProcedure.IdentifierName, memberName))
            {
                return enclosingProcedure;
            }

            return null;
        }

        public Declaration OnUndeclaredVariable(Declaration enclosingProcedure, string identifierName, ParserRuleContext context)
        {
            var annotations = _annotationService.FindAnnotations(enclosingProcedure.QualifiedName.QualifiedModuleName, context.Start.Line);
            var undeclaredLocal =
                new Declaration(
                    new QualifiedMemberName(enclosingProcedure.QualifiedName.QualifiedModuleName, identifierName),
                    enclosingProcedure,
                    enclosingProcedure,
                    "Variant",
                    string.Empty,
                    false,
                    false,
                    Accessibility.Implicit,
                    DeclarationType.Variable,
                    context,
                    null,
                    context.GetSelection(),
                    false,
                    null,
                    true,
                    annotations,
                    null,
                    true);

            var hasUndeclared = _newUndeclared.ContainsKey(enclosingProcedure.QualifiedName);
            if (hasUndeclared)
            {
                ConcurrentBag<Declaration> undeclared;
                while (!_newUndeclared.TryGetValue(enclosingProcedure.QualifiedName, out undeclared))
                {
                    _newUndeclared.TryGetValue(enclosingProcedure.QualifiedName, out undeclared);
                }
                var inScopeUndeclared = undeclared.FirstOrDefault(d => d.IdentifierName == identifierName);
                if (inScopeUndeclared != null)
                {
                    return inScopeUndeclared;
                }
                undeclared.Add(undeclaredLocal);
            }
            else
            {
                _newUndeclared.TryAdd(enclosingProcedure.QualifiedName, new ConcurrentBag<Declaration> { undeclaredLocal });
            }
            return undeclaredLocal;
        }


        public void AddUnboundContext(Declaration parentDeclaration, VBAParser.LExpressionContext context, IBoundExpression withExpression)
        {
            
            //The only forms we care about right now are MemberAccessExprContext or WithMemberAccessExprContext.
            if (!(context is VBAParser.MemberAccessExprContext) && !(context is VBAParser.WithMemberAccessExprContext))
            {
                return;
            }

            var identifier = context.GetChild<VBAParser.UnrestrictedIdentifierContext>(0);
            var annotations = _annotationService.FindAnnotations(parentDeclaration.QualifiedName.QualifiedModuleName, context.Start.Line);

            var declaration = new UnboundMemberDeclaration(parentDeclaration, identifier,
                (context is VBAParser.MemberAccessExprContext) ? (ParserRuleContext)context.children[0] : withExpression.Context, 
                annotations);

            _newUnresolved.Add(declaration);
        }

        public Declaration OnBracketedExpression(string expression, ParserRuleContext context)
        {
            var hostApp = FindProject(_hostApp == null ? "VBA" : _hostApp.ApplicationName);
            Debug.Assert(hostApp != null, "Host application project can't be null. Make sure VBA standard library is included if host is unknown.");

            var qualifiedName = hostApp.QualifiedName.QualifiedModuleName.QualifyMemberName(expression);

            if (_newUndeclared.TryGetValue(qualifiedName, out var undeclared))
            {
                return undeclared.SingleOrDefault();
            }

            var item = new Declaration(qualifiedName, hostApp, hostApp, Tokens.Variant, string.Empty, false, false, Accessibility.Global, DeclarationType.BracketedExpression, context, null, context.GetSelection(), true, null);
            _newUndeclared.TryAdd(qualifiedName, new ConcurrentBag<Declaration> { item });
            return item;
        }

        public Declaration FindMemberEnclosedProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, 
            string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m => m.DeclarationType.HasFlag(memberType)
                && (Declaration.GetModuleParent(m).DeclarationType == DeclarationType.ProceduralModule 
                    || memberType == DeclarationType.Enumeration 
                    || memberType == DeclarationType.EnumerationMember)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && !callingModule.Equals(Declaration.GetModuleParent(m)))
                .ToList();
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        private static bool IsInstanceSensitive(DeclarationType memberType)
        {
            return memberType.HasFlag(DeclarationType.Procedure)
                || memberType.HasFlag(DeclarationType.Function) 
                || memberType.HasFlag(DeclarationType.Variable)
                || memberType.HasFlag(DeclarationType.Constant);
        }

        public Declaration FindMemberEnclosedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, 
            string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && Declaration.GetProjectParent(m).Equals(callingProject)
                            && memberModule.Equals(Declaration.GetModuleParent(m)))
                .ToList();

            var match = memberMatches.FirstOrDefault(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            if (match != null)
            {
                return match;
            }

            return ClassModuleDeclaration
                .GetSupertypes(memberModule)
                .Select(supertype => 
                    FindMemberEnclosedProjectInModule(callingProject, callingModule, callingParent, supertype, memberName, memberType))
                .FirstOrDefault(supertypeMember => supertypeMember != null);
        }

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var isInstanceSensitive = IsInstanceSensitive(memberType);
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName,
                p => (!isInstanceSensitive || Declaration.GetModuleParent(p) == null ||
                      Declaration.GetModuleParent(p).DeclarationType != DeclarationType.ClassModule) &&
                     p.DeclarationType.HasFlag(memberType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, DeclarationType moduleType, 
            string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName,
                p => p.DeclarationType.HasFlag(memberType) &&
                     (Declaration.GetModuleParent(p) == null ||
                      Declaration.GetModuleParent(p).DeclarationType == moduleType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInGlobalClassModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, 
            string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(
                callingProject, 
                memberName, 
                p => p.DeclarationType.HasFlag(memberType) 
                    && (Declaration.GetModuleParent(p) == null 
                        || Declaration.GetModuleParent(p).DeclarationType == DeclarationType.ClassModule) 
                    && ((ClassModuleDeclaration)Declaration.GetModuleParent(p)).IsGlobalClassModule);
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, 
            string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(
                callingProject, 
                memberName, 
                p => p.DeclarationType.HasFlag(memberType) 
                    && memberModule.Equals(Declaration.GetModuleParent(p)
                ));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            if (match != null)
            {
                return match;
            }
            return ClassModuleDeclaration
                .GetSupertypes(memberModule)
                .Select(supertype => 
                    FindMemberReferencedProjectInModule(callingProject, callingModule, callingParent, supertype, memberName, memberType))
                .FirstOrDefault(supertypeMember => supertypeMember != null);
        }

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration referencedProject, 
            string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(
                callingProject, 
                memberName, 
                p => p.DeclarationType.HasFlag(memberType) 
                    && referencedProject.Equals(Declaration.GetProjectParent(p)
                ));
            return memberMatches.FirstOrDefault(m => 
                    AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
        }

        private Declaration FindInReferencedProjectByPriority(Declaration enclosingProject, string name, Func<Declaration, bool> predicate)
        {
            return FindAllInReferencedProjectByPriority(enclosingProject, name, predicate).FirstOrDefault();
        }

        private IEnumerable<Declaration> FindAllInReferencedProjectByPriority(Declaration enclosingProject, string name, Func<Declaration, bool> predicate)
        {
            var interprojectMatches = MatchName(name).Where(predicate).ToList();
            var projectReferences = ((ProjectDeclaration)enclosingProject).ProjectReferences;
            if (interprojectMatches.Count == 0)
            {
                yield break;
            }
            foreach (var projectReference in projectReferences)
            {
                var match = interprojectMatches.FirstOrDefault(interprojectMatch => interprojectMatch.ProjectId == projectReference.ReferencedProjectId);
                if (match != null)
                {
                    yield return match;
                }
            }
        }

        private IEnumerable<Declaration> FindAllFormControlHandlers()
        {
            var controls = DeclarationsWithType(DeclarationType.Control);
            var handlerNames = BuiltInDeclarations(DeclarationType.Event)
                .SelectMany(e => controls.Select(c => c.IdentifierName + "_" + e.IdentifierName))
                .ToHashSet();
            var handlers = UserDeclarations(DeclarationType.Procedure)
                .Where(procedure => handlerNames.Contains(procedure.IdentifierName));
            return handlers;
        }

        private List<Declaration> FindAllEventHandlers()
        {
            var handlerNames = BuiltInDeclarations(DeclarationType.Event)
                .SelectMany(e =>
                {
                    var parentModuleSubtypes = ((ClassModuleDeclaration)e.ParentDeclaration).Subtypes.ToList();
                    return parentModuleSubtypes.Any()
                        ? parentModuleSubtypes.Select(v => v.IdentifierName + "_" + e.IdentifierName)
                        : new[] { e.ParentDeclaration.IdentifierName + "_" + e.IdentifierName };
                })
                .ToHashSet();

            var handlers = DeclarationsWithType(DeclarationType.Procedure)
                .Where(item =>
                // class module built-in events
                (item.ParentDeclaration.DeclarationType == DeclarationType.ClassModule && (
                     item.IdentifierName.Equals("Class_Initialize", StringComparison.InvariantCultureIgnoreCase) ||
                     item.IdentifierName.Equals("Class_Terminate", StringComparison.InvariantCultureIgnoreCase))) ||
                // standard module built-in handlers (Excel specific):
                (_hostApp != null &&
                 _hostApp.ApplicationName.Equals("Excel", StringComparison.InvariantCultureIgnoreCase) &&
                 item.ParentDeclaration.DeclarationType == DeclarationType.ProceduralModule && (
                     item.IdentifierName.Equals("auto_open", StringComparison.InvariantCultureIgnoreCase) ||
                     item.IdentifierName.Equals("auto_close", StringComparison.InvariantCultureIgnoreCase))))
                .Concat(
                    UserDeclarations(DeclarationType.Procedure)
                        .Where(item => handlerNames.Contains(item.IdentifierName))
                )
                .Concat(_handlersByWithEventsField.Value.AllValues())
                .Concat(FindAllFormControlHandlers());
            return handlers.ToList();
        }

        /// <summary>
        /// Finds declarations that would be in conflict with the target declaration if renamed.
        /// </summary>
        /// <returns>Zero or more declarations that would be in conflict if the target declaration is renamed.</returns>
        public IEnumerable<Declaration> FindNewDeclarationNameConflicts(string newName, Declaration renameTarget)
        {
            if (newName.Equals(renameTarget.IdentifierName))
            {
                return Enumerable.Empty<Declaration>();
            }

            var identifierMatches = MatchName(newName);
            if (!identifierMatches.Any())
            {
                return Enumerable.Empty<Declaration>();
            }

            if (IsEnumOrUDTMemberDeclaration(renameTarget)) 
            {
                return identifierMatches.Where(idm =>
                    IsEnumOrUDTMemberDeclaration(idm) && idm.ParentDeclaration == renameTarget.ParentDeclaration);
            }

            identifierMatches = identifierMatches.Where(nc => !IsEnumOrUDTMemberDeclaration(nc));
            var referenceConflicts = identifierMatches.Where(idm =>
                renameTarget.References
                    .Any(renameTargetRef => renameTargetRef.ParentScoping == idm.ParentDeclaration
                        || renameTarget.ParentDeclaration.DeclarationType != DeclarationType.ClassModule
                            && idm == renameTargetRef.ParentScoping
                            && !UsesScopeResolution(renameTargetRef.Context.Parent)
                        || idm.References
                            .Any(idmRef => idmRef.ParentScoping == renameTargetRef.ParentScoping
                                && !UsesScopeResolution(renameTargetRef.Context.Parent)))
                || idm.DeclarationType.HasFlag(DeclarationType.Variable)
                    && idm.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module)
                    && renameTarget.References.Any(renameTargetRef => renameTargetRef.QualifiedModuleName == idm.ParentDeclaration.QualifiedModuleName));

            if (referenceConflicts.Any())
            {
                return referenceConflicts;
            }

            var renameTargetModule = Declaration.GetModuleParent(renameTarget);
            var declarationConflicts = identifierMatches.Where(idm =>
                renameTarget == idm.ParentDeclaration
                || AccessibilityCheck.IsAccessible(
                    Declaration.GetProjectParent(renameTarget),
                    renameTargetModule,
                    renameTarget.ParentDeclaration,
                    idm)
                    && IsConflictingMember(renameTarget, renameTargetModule, idm));

            return declarationConflicts;
        }

        private bool IsEnumOrUDTMemberDeclaration(Declaration candidate)
        {
            return candidate.DeclarationType == DeclarationType.EnumerationMember
                       || candidate.DeclarationType == DeclarationType.UserDefinedTypeMember;
        }

        private bool IsConflictingMember(Declaration renameTarget, Declaration renameTargetModule, Declaration candidate)
        {
            var candidateModule = Declaration.GetModuleParent(candidate);
            return renameTargetModule == candidateModule
             || renameTargetModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule)
                && candidate.Accessibility != Accessibility.Private
                && candidateModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule);
        }

        private bool UsesScopeResolution(RuleContext ruleContext)
        {
            return (ruleContext is VBAParser.WithMemberAccessExprContext)
                || (ruleContext is VBAParser.MemberAccessExprContext);
        }

        /// <summary>
        /// Creates a dictionary of identifier references, keyed by module.
        /// </summary>
        public IReadOnlyDictionary<QualifiedModuleName, IReadOnlyList<IdentifierReference>> IdentifierReferences()
        {
            return _referencesByModule;
        }

        /// <summary>
        /// Gets all identifier references in the specified module.
        /// </summary>
        public IEnumerable<IdentifierReference> IdentifierReferences(QualifiedModuleName module)
        {
            return _referencesByModule.TryGetValue(module, out var value)
                ? value
                : Enumerable.Empty<IdentifierReference>();
        }

        /// <summary>
        /// Gets all identifier references in the specified member.
        /// </summary>
        public IEnumerable<IdentifierReference> IdentifierReferences(QualifiedMemberName member)
        {
            return _referencesByMember.TryGetValue(member, out List<IdentifierReference> value)
                ? value
                : Enumerable.Empty<IdentifierReference>();
        }

        /// <summary>
        /// Gets all identifier references.
        /// </summary>
        public IEnumerable<IdentifierReference> AllIdentifierReferences()
        {
            return _referencesByModule.Values.SelectMany(list => list);
        }
    }
}
