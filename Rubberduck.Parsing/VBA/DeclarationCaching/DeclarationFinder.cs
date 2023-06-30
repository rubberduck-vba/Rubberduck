using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using NLog;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.DeclarationCaching
{
    public class DeclarationFinder
    {
        private static readonly SquareBracketedNameComparer NameComparer = new SquareBracketedNameComparer();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IHostApplication _hostApp;
        private IDictionary<string, List<Declaration>> _declarationsByName;
        private IDictionary<QualifiedModuleName, IDictionary<DeclarationType, List<Declaration>>> _declarations;

        private readonly IReadOnlyDictionary<QualifiedModuleName, IFailedResolutionStore> _failedResolutionStores;
        private readonly ConcurrentDictionary<QualifiedModuleName, IMutableFailedResolutionStore> _newFailedResolutionStores;
        private readonly ConcurrentDictionary<(QualifiedMemberName memberName, DeclarationType declarationType), ConcurrentBag<Declaration>> _newUndeclared;

        private IDictionary<QualifiedModuleName,IDictionary<int, List<IParseTreeAnnotation>>> _annotations;
        private readonly IReadOnlyDictionary<QualifiedModuleName, LogicalLineStore> _logicalLines;
        private IDictionary<Declaration, List<ParameterDeclaration>> _parametersByParent;
        private IDictionary<DeclarationType, List<Declaration>> _userDeclarationsByType;
       
        private IReadOnlyList<IdentifierReference> _identifierReferences;
        private IDictionary<QualifiedSelection, List<IdentifierReference>> _referencesBySelection;
        private IReadOnlyDictionary<QualifiedModuleName, IReadOnlyList<IdentifierReference>> _referencesByModule;
        private IReadOnlyDictionary<string, IReadOnlyList<IdentifierReference>> _referencesByProjectId;
        private IDictionary<QualifiedMemberName, List<IdentifierReference>> _referencesByMember;

        private Lazy<IDictionary<DeclarationType, List<Declaration>>> _builtInDeclarationsByType;
        private Lazy<IDictionary<Declaration, List<ModuleBodyElementDeclaration>>> _handlersByWithEventsField;

        private Lazy<IDictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>>> _implementingMembers;
        private Lazy<IDictionary<VBAParser.ImplementsStmtContext, List<ModuleBodyElementDeclaration>>> _membersByImplementsContext;
        private Lazy<IDictionary<ClassModuleDeclaration, List<Declaration>>> _interfaceMembers;
        private Lazy<IDictionary<ClassModuleDeclaration, List<ClassModuleDeclaration>>> _interfaceImplementations;
        private Lazy<IDictionary<IInterfaceExposable, List<ModuleBodyElementDeclaration>>> _implementationsByMember;

        private Lazy<ICollection<Declaration>> _nonBaseAsType; 
        private Lazy<ICollection<Declaration>> _eventHandlers;
        private Lazy<ICollection<Declaration>> _controlEventHandlers;
        private Lazy<ICollection<Declaration>> _formEventHandlers;
        private Lazy<ICollection<Declaration>> _projects;
        private Lazy<ICollection<Declaration>> _classes;
        
        private static QualifiedSelection GetGroupingKey(Declaration declaration)
        {
            // we want the procedures' whole body, not just their identifier:
            return declaration.DeclarationType.HasFlag(DeclarationType.Member)
                ? new QualifiedSelection(
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.Context.GetSelection())
                : declaration.QualifiedSelection;
        }
        
        public DeclarationFinder(IReadOnlyList<Declaration> declarations,
            IEnumerable<IParseTreeAnnotation> annotations,
            IReadOnlyDictionary<QualifiedModuleName, LogicalLineStore> logicalLines,
            IReadOnlyDictionary<QualifiedModuleName, IFailedResolutionStore> failedResolutionStores,
            IHostApplication hostApp = null)
        {
            _hostApp = hostApp;
            _failedResolutionStores = failedResolutionStores;
            _logicalLines = logicalLines;

            _newFailedResolutionStores = new ConcurrentDictionary<QualifiedModuleName, IMutableFailedResolutionStore>();
            _newUndeclared = new ConcurrentDictionary<(QualifiedMemberName memberName, DeclarationType declarationType), ConcurrentBag<Declaration>>();

            var collectionConstructionActions = CollectionConstructionActions(declarations, annotations);
            ExecuteCollectionConstructionActions(collectionConstructionActions);

            //Temporal coupling: the initializers of the lazy collections use the regular collections filled above.
            InitializeLazyCollections();
        }

        protected virtual void ExecuteCollectionConstructionActions(List<Action> collectionConstructionActions)
        {
            collectionConstructionActions.ForEach(action => action.Invoke());
        }

        private List<Action> CollectionConstructionActions(IReadOnlyList<Declaration> declarations, IEnumerable<IParseTreeAnnotation> annotations)
        {
            var actions = new List<Action>
            {
                () =>
                    _annotations = annotations
                        .GroupBy(annotation => annotation.QualifiedSelection.QualifiedName)
                        .SelectMany(grp1 =>  grp1.GroupBy(annotation => annotation.AnnotatedLine.GetValueOrDefault(-1)), (grp1, grp2) => (grp1, grp2))
                        .GroupBy(tpl => tpl.grp1.Key, tpl => tpl.grp2)
                        .ToDictionary(),
                () =>
                    _declarations = declarations
                        .GroupBy(item => item.QualifiedName.QualifiedModuleName)
                        .SelectMany(grp1 => grp1.GroupBy(declaration => declaration.DeclarationType), (grp1, grp2) => (grp1, grp2))
                        .GroupBy(tpl => tpl.grp1.Key, tpl => tpl.grp2)
                        .ToDictionary(),
                () =>
                    _declarationsByName = declarations
                        .GroupBy(declaration => declaration.IdentifierName.ToLowerInvariant())
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

                () => InitializeIdentifierDictionaries(declarations)
            };

            return actions;
        }

        private void InitializeIdentifierDictionaries(IReadOnlyList<Declaration> declarations)
        {
            _identifierReferences = declarations.SelectMany(declaration => declaration.References).ToList();

            _referencesBySelection = _identifierReferences
                .GroupBy(reference => new QualifiedSelection(reference.QualifiedModuleName, reference.Selection))
                .ToDictionary();

            _referencesByModule = _identifierReferences
                .GroupBy(reference => Declaration.GetModuleParent(reference.ParentScoping).QualifiedName.QualifiedModuleName)
                .ToReadonlyDictionary();

            _referencesByMember = _identifierReferences
                .GroupBy(reference => reference.ParentScoping.QualifiedName)
                .ToDictionary();

            _referencesByProjectId = _identifierReferences
                .GroupBy(reference => reference.Declaration.ProjectId)
                .ToReadonlyDictionary();
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

            _nonBaseAsType = new Lazy<ICollection<Declaration>>(() =>
                _declarations
                    .AllValues()
                    .Where(d => !string.IsNullOrWhiteSpace(d.AsTypeName)
                                    && !d.AsTypeIsBaseType
                                    && d.DeclarationType != DeclarationType.Project
                                    && d.DeclarationType != DeclarationType.ProceduralModule)
                    .ToList()
                , true);

            _eventHandlers = new Lazy<ICollection<Declaration>>(FindAllEventHandlers, true);
            _controlEventHandlers = new Lazy<ICollection<Declaration>>(FindAllFormControlHandlers, true);
            _formEventHandlers = new Lazy<ICollection<Declaration>>(FindAllFormEventHandlers, true);
            _projects = new Lazy<ICollection<Declaration>>(() => DeclarationsWithType(DeclarationType.Project).ToList(), true);
            _classes = new Lazy<ICollection<Declaration>>(() => DeclarationsWithType(DeclarationType.ClassModule).ToList(), true);
            _handlersByWithEventsField = new Lazy<IDictionary<Declaration, List<ModuleBodyElementDeclaration>>>(FindAllHandlersByWithEventField, true);

            _implementingMembers = new Lazy<IDictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>>>(FindAllImplementingMembers, true);
            _interfaceMembers = new Lazy<IDictionary<ClassModuleDeclaration, List<Declaration>>>(FindAllInterfaceMembersByModule, true);
            _membersByImplementsContext = new Lazy<IDictionary<VBAParser.ImplementsStmtContext, List<ModuleBodyElementDeclaration>>>(FindAllImplementingMembersByImplementsContext, true);
            _interfaceImplementations = new Lazy<IDictionary<ClassModuleDeclaration, List<ClassModuleDeclaration>>>(FindAllImplementionsByInterface, true);
            _implementationsByMember = new Lazy<IDictionary<IInterfaceExposable, List<ModuleBodyElementDeclaration>>>(FindAllImplementingMembersByMember, true);
        }

        private IDictionary<(VBAParser.ImplementsStmtContext Context, Declaration Implementor), List<ModuleBodyElementDeclaration>> FindAllImplementingMembers()
        {
            var implementsInstructions = UserDeclarations(DeclarationType.ClassModule)
                .Concat(UserDeclarations(DeclarationType.Document))
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
                .Concat(UserDeclarations(DeclarationType.Document))
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

        private IDictionary<ClassModuleDeclaration, List<Declaration>> FindAllInterfaceMembersByModule()
        {
            return UserDeclarations(DeclarationType.ClassModule)
                .Concat(UserDeclarations(DeclarationType.Document))
                .Concat(UserDeclarations(DeclarationType.UserForm))
                .Cast<ClassModuleDeclaration>()
                .Where(module => module.IsInterface)
                .ToDictionary(
                    module => module,
                    module => module.Members
                        .Where(member => member is IInterfaceExposable candidate && candidate.IsInterfaceMember)
                        .ToList());
        }

        private IDictionary<Declaration, List<ModuleBodyElementDeclaration>> FindAllHandlersByWithEventField()
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
                            Members(item.WithEventsField.ParentDeclaration.QualifiedName.QualifiedModuleName, DeclarationType.Procedure)
                                .Where(member => member.IdentifierName.Equals($"{item.WithEventsField.IdentifierName}_{evnt.IdentifierName}", StringComparison.InvariantCultureIgnoreCase)))
                            .OfType<ModuleBodyElementDeclaration>()
                    })
                    .ToDictionary(item => item.WithEventsField, item => item.Handlers.ToList());
            return handlersByWithEventsField;
        }

        //This does not need a lock because enumerators over a ConcurrentBag uses a snapshot.    
        public IEnumerable<Declaration> FreshUndeclared => _newUndeclared.AllValues();
        public IReadOnlyDictionary<QualifiedModuleName, IFailedResolutionStore> FreshFailedResolutionStores => _newFailedResolutionStores.ToDictionary(kvp => kvp.Key, kvp => (IFailedResolutionStore)new FailedResolutionStore(kvp.Value));

        public int? LogicalLine(QualifiedModuleName module, int physicalLine)
        {
            return _logicalLines.TryGetValue(module, out var lineStore)
                ? lineStore.LogicalLineNumber(physicalLine)
                : null;
        }

        public int? PhysicalStartLine(QualifiedModuleName module, int logicalLine)
        {
            return _logicalLines.TryGetValue(module, out var lineStore)
                ? lineStore.PhysicalStartLineNumber(logicalLine)
                : null;
        }

        public int? PhysicalEndLine(QualifiedModuleName module, int logicalLine)
        {
            return _logicalLines.TryGetValue(module, out var lineStore)
                ? lineStore.PhysicalEndLineNumber(logicalLine)
                : null;
        }

        public int? NumberOfLogicalLines(QualifiedModuleName module)
        {
            if (!_logicalLines.TryGetValue(module, out var lineStore))
            {
                return null;
            }
            return lineStore.NumberOfLogicalLines();
        }

        public int? StartOfContainingLogicalLine(QualifiedModuleName module, int physicalLine)
        {
            return _logicalLines.TryGetValue(module, out var lineStore)
                ? lineStore.StartOfContainingLogicalLine(physicalLine)
                : null;
        }

        public int? EndOfContainingLogicalLine(QualifiedModuleName module, int physicalLine)
        {
            return _logicalLines.TryGetValue(module, out var lineStore)
                ? lineStore.EndOfContainingLogicalLine(physicalLine)
                : null;
        }

        public IEnumerable<Declaration> Members(Declaration module)
        {
            return Members(module.QualifiedName.QualifiedModuleName);
        }

        public IEnumerable<Declaration> Members(Declaration module, DeclarationType declarationType)
        {
            return Members(module.QualifiedName.QualifiedModuleName, declarationType);
        }

        public IEnumerable<Declaration> Members(QualifiedModuleName module)
        {
            return _declarations.TryGetValue(module, out var members)
                    ? members.AllValues()
                    : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<Declaration> Members(QualifiedModuleName module, DeclarationType declarationType)
        {
            if (!_declarations.TryGetValue(module, out var membersByType))
            {
                return Enumerable.Empty<Declaration>();
            }

            return membersByType
                    .Where(item => item.Key.HasFlag(declarationType))
                    .SelectMany(item => item.Value);
        }

        public Declaration ModuleDeclaration(QualifiedModuleName module)
        {
            return Members(module, DeclarationType.Module).SingleOrDefault();
        }

        public IReadOnlyCollection<QualifiedModuleName> AllModules => _declarations.Keys.AsReadOnly();

        public IEnumerable<Declaration> AllDeclarations => _declarations.AllValues();

        public ICollection<Declaration> FindDeclarationsWithNonBaseAsType()
        {
            return _nonBaseAsType.Value;
        }
 
        public ICollection<Declaration> FindEventHandlers()
        {
            return _eventHandlers.Value;
        }

        public IEnumerable<ModuleBodyElementDeclaration> FindEventHandlers(Declaration eventDeclaration)
        {
            var withEventsDeclarations = FindWithEventFields(eventDeclaration);
            return withEventsDeclarations
                .Select(withEventsField => FindHandlersForWithEventsField(withEventsField).SingleOrDefault(handler => 
                    handler.IdentifierName.Equals($"{withEventsField.IdentifierName}_{eventDeclaration.IdentifierName}", StringComparison.InvariantCultureIgnoreCase)));
        }

        public ICollection<Declaration> FindFormControlEventHandlers()
        {
            return _controlEventHandlers.Value;
        }

        public IEnumerable<Declaration> FindFormControlEventHandlers(Declaration control)
        {
            return _eventHandlers.Value
                .Where(handlers=> handlers.ParentScope.Equals(control.ParentScope, StringComparison.InvariantCultureIgnoreCase) && 
                    handlers.IdentifierName.StartsWith(control.IdentifierName + "_", StringComparison.InvariantCultureIgnoreCase));
        }

        public ICollection<Declaration> FindFormEventHandlers()
        {
            return _formEventHandlers.Value;
        }

        public IEnumerable<Declaration> Classes => _classes.Value;
        public IEnumerable<Declaration> Projects => _projects.Value;

        /// <summary>
        /// Gets the <see cref="ProjectDeclaration"/> object for specified referenced project/library.
        /// </summary>
        /// <param name="name">The identifier name of the project declaration to find.</param>
        /// <param name="result">The <see cref="ProjectDeclaration"/> result, if found; null otherwise.</param>
        /// <param name="includeUserDefined">True to include user-defined projects in the search; false by default.</param>
        public bool TryFindProjectDeclaration(string name, out Declaration result, bool includeUserDefined = false)
        {
            result = _projects.Value.FirstOrDefault(project => project.IdentifierName.Equals(name, StringComparison.InvariantCultureIgnoreCase) && project.IsUserDefined == includeUserDefined);
            return result != null;
        }

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

        public IEnumerable<ModuleBodyElementDeclaration> FindHandlersForWithEventsField(Declaration field)
        {
            return _handlersByWithEventsField.Value.TryGetValue(field, out var result) 
                ? result 
                : Enumerable.Empty<ModuleBodyElementDeclaration>();
        }

        public IEnumerable<Declaration> FindWithEventFields()
        {
            return _handlersByWithEventsField.Value.Keys;
        }

        public IEnumerable<Declaration> FindWithEventFields(Declaration eventDeclaration)
        {
            return FindWithEventFields().Where(withEventField => withEventField.AsTypeName == eventDeclaration.ComponentName);
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
                    .Where(reference => reference.QualifiedModuleName.Equals(selection.QualifiedName))
                    .Select(reference => reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>())
                    .Where(context => context != null)
                    .Select(context => context.GetSelection())
                    .Any(contextSelection => contextSelection.Contains(selection.Selection) 
                                             || selection.Selection.Contains(contextSelection)));
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
            var lookup = _interfaceImplementations.Value;
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

        public ParameterDeclaration FindParameter(Declaration parameterizedMember, string parameterName)
        {
            return _parametersByParent.TryGetValue(parameterizedMember, out List<ParameterDeclaration> parameters) 
                ? parameters.SingleOrDefault(parameter => parameter.IdentifierName == parameterName) 
                : null;
        }

        /// <summary>
        /// Returns the parameter at index parameterIndex (0-based)
        /// </summary>
        public ParameterDeclaration FindParameter(Declaration parameterizedMember, int parameterIndex)
        {
            return parameterIndex >= 0
                && _parametersByParent.TryGetValue(parameterizedMember, out List<ParameterDeclaration> parameters)
                && parameterIndex < parameters.Count
                ? parameters[parameterIndex]
                : null;
        }

        public IEnumerable<ParameterDeclaration> Parameters(Declaration parameterizedMember)
        {
            return _parametersByParent.TryGetValue(parameterizedMember, out List<ParameterDeclaration> result)
                ? result
                : Enumerable.Empty<ParameterDeclaration>();
        }

        public IEnumerable<Declaration> FindMemberMatches(Declaration parent, string memberName)
        {
            return Members(parent.QualifiedName.QualifiedModuleName, DeclarationType.Member)
                .Where(member => member.IdentifierName.Equals(memberName));
        }

        public IEnumerable<IParseTreeAnnotation> FindAnnotations(QualifiedModuleName module, int annotatedLine)
        {
            if (!_annotations.TryGetValue(module, out var annotationsByLineInModule))
            {
                return Enumerable.Empty<IParseTreeAnnotation>();
            }

            var firstLineOfAnnotatedLogicalLine = StartOfContainingLogicalLine(module, annotatedLine);
            if (!firstLineOfAnnotatedLogicalLine.HasValue)
            {
                return Enumerable.Empty<IParseTreeAnnotation>();
            }

            return annotationsByLineInModule.TryGetValue(firstLineOfAnnotatedLogicalLine.Value, out var result) 
                ? result 
                : Enumerable.Empty<IParseTreeAnnotation>();
        }

        public IEnumerable<IParseTreeAnnotation> FindAnnotations(QualifiedModuleName module)
        {
            return _annotations.TryGetValue(module, out var annotationsByLineInModule)
                ? annotationsByLineInModule.AllValues()
                : Enumerable.Empty<IParseTreeAnnotation>();
        }

        public IEnumerable<IParseTreeAnnotation> FindAnnotations(QualifiedModuleName module, int annotatedLine, Type annotationType)
        {
            return FindAnnotations(module, annotatedLine).Where(pta => pta.Annotation.GetType() == annotationType);
        }

        public IEnumerable<IParseTreeAnnotation> FindAnnotations<T>(QualifiedModuleName module, int annotatedLine) 
             where T : IAnnotation
        {
            return FindAnnotations(module, annotatedLine, typeof(T));
        }

        public IEnumerable<IParseTreeAnnotation> FindAnnotations(QualifiedModuleName module, int annotatedLine, AnnotationTarget target)
        {
            return FindAnnotations(module, annotatedLine)
                .Where(annot => annot.Annotation.Target.HasFlag(target));
        }

        public bool IsMatch(string declarationName, string potentialMatchName)
        {
            return string.Equals(declarationName, potentialMatchName, StringComparison.InvariantCultureIgnoreCase);
        }

        private IEnumerable<Declaration> FindEvents(Declaration module)
        {
            if (module is null)
            {
                return Enumerable.Empty<Declaration>();
            }

            return Members(module.QualifiedName.QualifiedModuleName, DeclarationType.Event);
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

        public ParameterDeclaration FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(VBAParser.ArgumentContext argument, Declaration enclosingProcedure)
        {
            return FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argument, enclosingProcedure.QualifiedModuleName);
        }

        public ParameterDeclaration FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(VBAParser.ArgumentContext argument, QualifiedModuleName module)
        {
            //todo: Rename after making it work for more general cases.
            var missingArgument = argument.missingArgument();
            var argumentExpression = argument.GetDescendent<VBAParser.ArgumentExpressionContext>();
            if ((missingArgument == null && argumentExpression  == null)
                || argumentExpression?.GetDescendent<VBAParser.ParenthesizedExprContext>() != null 
                || argumentExpression?.BYVAL() != null)
            {
                // not a simple argument, or argument is parenthesized and thus passed ByVal
                return null;
            }

            var callingNonDefaultMember = CallingNonDefaultMember((ParserRuleContext)argumentExpression ?? missingArgument, module);
            if (callingNonDefaultMember == null)
            {
                // Either we could not resolve the call or there is a default member call involved. 
                return null;
            }

            var parameters = Parameters(callingNonDefaultMember);
            ParameterDeclaration parameter;

            var namedArg = argumentExpression.GetAncestor<VBAParser.NamedArgumentContext>();
            if (namedArg != null)
            {
                // argument is named: we're lucky
                var parameterName = namedArg.unrestrictedIdentifier().GetText();
                parameter = parameters.SingleOrDefault(p => p.IdentifierName == parameterName);
            }
            else
            {
                // argument is positional: work out its index
                var argumentList = ((ParserRuleContext)argumentExpression ?? missingArgument).GetAncestor<VBAParser.ArgumentListContext>();
                var arguments = argumentList.children.Where(t => t is VBAParser.ArgumentContext).ToArray();
                var selection = argumentExpression?.GetSelection() ?? missingArgument.GetSelection();
                
                var indexedArgs = arguments.Select((arg, index) => (arg: arg as ParserRuleContext, index))
                    .Select(e => (arg: e.arg, e.index, selection:e.arg.GetSelection()))
                    .ToList();
                var indexedArg = indexedArgs.SingleOrDefault(item => item.selection.Contains(selection));
                if (indexedArg.arg == null)
                {
                    return null;
                }
                parameter = parameters
                    .Select((param, index) => (param, index))
                    .SingleOrDefault(item => item.index == indexedArg.index).param;
            }

            return parameter;
        }

        public ModuleBodyElementDeclaration FindInvokedMemberFromArgumentContext(VBAParser.ArgumentContext argument, QualifiedModuleName module)
        {
            var expression = (ParserRuleContext)argument.GetDescendent<VBAParser.ArgumentExpressionContext>() 
                ?? argument.GetDescendent<VBAParser.MissingArgumentContext>();
            return expression != null
                ? CallingNonDefaultMember(expression, module)
                : null;
        }

        private ModuleBodyElementDeclaration CallingNonDefaultMember(ParserRuleContext argumentExpressionOrMissingArgument, QualifiedModuleName module)
        {
            //todo: Make this work for default member calls.

            var argumentList = argumentExpressionOrMissingArgument.GetAncestor<VBAParser.ArgumentListContext>();
            var cannotHaveDefaultMemberCall = false;

            ParserRuleContext callingExpression;
            switch (argumentList?.Parent)
            {
                case VBAParser.CallStmtContext callStmt:
                    cannotHaveDefaultMemberCall = true;
                    callingExpression = callStmt.lExpression();
                    break;
                case VBAParser.IndexExprContext indexExpr:
                    callingExpression = indexExpr.lExpression();
                    break;
                case VBAParser.WhitespaceIndexExprContext indexExpr:
                    callingExpression = indexExpr.lExpression();
                    break;
                default:
                    return null;
            }

            VBAParser.IdentifierContext callingIdentifier;
            if (cannotHaveDefaultMemberCall)
            {
                callingIdentifier = callingExpression
                    .GetDescendents<VBAParser.IdentifierContext>()
                    .LastOrDefault();
            }
            else
            {
                switch (callingExpression)
                {
                    case VBAParser.SimpleNameExprContext simpleName:
                        callingIdentifier = simpleName.identifier();
                        break;
                    case VBAParser.MemberAccessExprContext memberAccess:
                        callingIdentifier = memberAccess
                            .GetDescendents<VBAParser.IdentifierContext>()
                            .LastOrDefault();
                        break;
                    case VBAParser.WithMemberAccessExprContext memberAccess:
                        callingIdentifier = memberAccess
                            .GetDescendents<VBAParser.IdentifierContext>()
                            .LastOrDefault();
                        break;
                    default:
                        //This is only possible in case of a default member access.
                        return null;
                }
            }

            if (callingIdentifier == null)
            {
                return null;
            }

            var referencedMember = IdentifierReferences(callingIdentifier, module)
                .Select(reference => reference.Declaration)
                .OfType<ModuleBodyElementDeclaration>()
                .FirstOrDefault();

            return referencedMember;
        }

        public ParameterDeclaration FindParameterFromSimpleEventArgumentNotPassedByValExplicitly(VBAParser.EventArgumentContext eventArgument, QualifiedModuleName module)
        {
            if (eventArgument == null
                || eventArgument.GetDescendent<VBAParser.ParenthesizedExprContext>() != null
                || eventArgument.BYVAL() != null)
            {
                // not a simple argument, or argument is parenthesized and thus passed ByVal
                return null;
            }

            var raisedEvent = RaisedEvent(eventArgument, module);
            if (raisedEvent == null)
            {
                return null;
            }

            var parameters = Parameters(raisedEvent);

            // event arguments are always positional: work out the index
            var argumentList = eventArgument.GetAncestor<VBAParser.EventArgumentListContext>();
            var arguments = argumentList.eventArgument();

            var parameterIndex = arguments
                .Select((arg, index) => arg == eventArgument ? (arg, index) : (null, -1))
                .SingleOrDefault(tpl => tpl.arg != null).index;

            var parameter = parameters
                .OrderBy(p => p.Selection)
                .Select((param, index) => (param, index))
                .SingleOrDefault(tpl => tpl.index == parameterIndex).param;

            return parameter;
        }

        private EventDeclaration RaisedEvent(VBAParser.EventArgumentContext argument, QualifiedModuleName module)
        {
            var raiseEventContext = argument.GetAncestor<VBAParser.RaiseEventStmtContext>();
            var eventIdentifier = raiseEventContext.identifier();

            var referencedMember = IdentifierReferences(eventIdentifier, module)
                .Select(reference => reference.Declaration)
                .OfType<EventDeclaration>()
                .FirstOrDefault();

            return referencedMember;
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
                m is ClassModuleDeclaration classModule 
                && classModule.HasDefaultInstanceVariable
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
                p => p is ClassModuleDeclaration classModule &&
                     classModule.HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, 
            string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName,
                p => referencedProject.Equals(Declaration.GetProjectParent(p))
                    && p is ClassModuleDeclaration classModule
                    && classModule.HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberWithParent(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration parent, 
            string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var parentClass = parent as ClassModuleDeclaration;
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && (parent.Equals(m.ParentDeclaration) || (parentClass?.Supertypes.Any(t => t.Equals(m.ParentDeclaration)) ?? false)))
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
            if (callingModule is ClassModuleDeclaration callingClass)
            {
                foreach (var supertype in callingClass.Supertypes)
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

        //TODO: ove this out of the DeclarationFinder into a class only responsible for reference resolution.
        public Declaration OnUndeclaredVariable(Declaration enclosingProcedure, string identifierName, ParserRuleContext context)
        {
            var annotations = FindAnnotations(enclosingProcedure.QualifiedName.QualifiedModuleName, context.Start.Line,AnnotationTarget.Identifier);
            var isReDimVariable = IsContainedInReDimedArrayName(context);

            Declaration undeclaredLocal;
            if (IsContainedInReDimedArrayName(context))
            {
                var asTypeClause = AsTypeClauseForReDimDeclaredArray(context);
                var typeHint = TypeHintForReDimDeclaredArray(context);
                var asTypeName = AsTypeNameForReDimDeclaredArray(asTypeClause, typeHint);

                undeclaredLocal = new Declaration(
                    new QualifiedMemberName(enclosingProcedure.QualifiedName.QualifiedModuleName, identifierName),
                    enclosingProcedure,
                    enclosingProcedure,
                    asTypeName,
                    typeHint,
                    false,
                    false,
                    Accessibility.Implicit,
                    DeclarationType.Variable,
                    context,
                    null,
                    context.GetSelection(),
                    true,
                    asTypeClause,
                    true,
                    annotations,
                    null,
                    false);

                ResolveTypeForReDimDeclaredArray(undeclaredLocal);
            }
            else 
            {
                undeclaredLocal = new Declaration(
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
            }

            var enclosingScope = (enclosingProcedure.QualifiedName, enclosingProcedure.DeclarationType);
            var hasUndeclared = _newUndeclared.ContainsKey(enclosingScope);
            if (hasUndeclared)
            {
                ConcurrentBag<Declaration> undeclared;
                while (!_newUndeclared.TryGetValue(enclosingScope, out undeclared))
                {
                    _newUndeclared.TryGetValue(enclosingScope, out undeclared);
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
                _newUndeclared.TryAdd(enclosingScope, new ConcurrentBag<Declaration> { undeclaredLocal });
            }
            return undeclaredLocal;
        }

        private static bool IsContainedInReDimedArrayName(ParserRuleContext context)
        {
            return ContainingReDimContext(context) != null;
        }

        private static VBAParser.RedimVariableDeclarationContext ContainingReDimContext(ParserRuleContext context)
        {
            var enclosingReDimContextCandidate = context.GetAncestor<VBAParser.RedimVariableDeclarationContext>();
            return enclosingReDimContextCandidate == null
                   || !enclosingReDimContextCandidate.expression().GetSelection().Contains(context.GetSelection())
                   ? null
                   : enclosingReDimContextCandidate;
        }

        private VBAParser.AsTypeClauseContext AsTypeClauseForReDimDeclaredArray(ParserRuleContext context)
        {
            return ContainingReDimContext(context)?.asTypeClause();
        }

        private string TypeHintForReDimDeclaredArray(ParserRuleContext context)
        {
            if (context is VBAParser.SimpleNameExprContext simpleNameContext)
            {
                var identifier = simpleNameContext.identifier();
                return identifier != null
                    ? Identifier.GetTypeHintValue(identifier)
                    : null;
            }

            return null;
        }

        private static string AsTypeNameForReDimDeclaredArray(VBAParser.AsTypeClauseContext asTypeClause, string typeHint)
        {
            return typeHint == null
                ? asTypeClause == null
                    ? Tokens.Variant
                    : asTypeClause.type().GetText()
                : SymbolList.TypeHintToTypeName[typeHint];
        }

        //note: This is copied from the TypeAnnotationPass.
        //TODO: Extract common logic from TypeAnnotationPass and inject that. (Requires extracting the calling code out of the DeclarationFinder.)
        private void ResolveTypeForReDimDeclaredArray(Declaration declaration)
        {
            if(string.IsNullOrWhiteSpace(declaration.AsTypeName)
                || declaration.AsTypeIsBaseType)
            {
                return;
            }

            string typeExpression;
            if (declaration.AsTypeContext != null && declaration.AsTypeContext.type().complexType() != null)
            {
                var typeContext = declaration.AsTypeContext;
                typeExpression = typeContext.type().complexType().GetText();
            }
            else if (!string.IsNullOrWhiteSpace(declaration.AsTypeNameWithoutArrayDesignator) && !SymbolList.BaseTypes.Contains(declaration.AsTypeNameWithoutArrayDesignator.ToUpperInvariant()))
            {
                typeExpression = declaration.AsTypeNameWithoutArrayDesignator;
            }
            else
            {
                return;
            }

            var module = Declaration.GetModuleParent(declaration);
            if (module == null)
            {
                Logger.Warn("Type annotation failed for {0} because module parent is missing.", typeExpression);
                return;
            }

            var (bindingService, expressionParser) = TypeAnnotationServices();

            var expressionContext = expressionParser.Parse(typeExpression.Trim());
            var boundExpression = bindingService.ResolveType(module, declaration.ParentDeclaration, expressionContext);
            if (boundExpression.Classification != ExpressionClassification.ResolutionFailed)
            {
                declaration.AsTypeDeclaration = boundExpression.ReferencedDeclaration;
            }
            else
            {
                const string IGNORE_THIS = "DISPATCH";
                if (typeExpression != IGNORE_THIS)
                {
                    Logger.Warn("Failed to resolve type {0}", typeExpression);
                }
            }
        }

        private (BindingService bindingservice, VBAExpressionParser expressionParser) TypeAnnotationServices()
        {
            var typeBindingContext = new TypeBindingContext(this);
            var procedurePointerBindingContext = new ProcedurePointerBindingContext(this);
            var bindingService = new BindingService(
                this,
                new DefaultBindingContext(this, typeBindingContext, procedurePointerBindingContext),
                typeBindingContext,
                procedurePointerBindingContext);
            var expressionParser = new VBAExpressionParser();
            return (bindingService, expressionParser);
        }

        public void AddUnboundContext(Declaration parentDeclaration, VBAParser.LExpressionContext context, IBoundExpression withExpression)
        {
            
            //The only forms we care about right now are MemberAccessExprContext or WithMemberAccessExprContext.
            //For WithMemberAccessExpressions we can only save an unbout member if the withExpression is not null.
            if (!(context is VBAParser.MemberAccessExprContext) && !(context is VBAParser.WithMemberAccessExprContext && withExpression != null))
            {
                return;
            }

            var identifier = context.GetChild<VBAParser.UnrestrictedIdentifierContext>(0);
            var annotations = FindAnnotations(parentDeclaration.QualifiedName.QualifiedModuleName, context.Start.Line, AnnotationTarget.Identifier);

            var declaration = new UnboundMemberDeclaration(parentDeclaration, identifier,
                (context is VBAParser.MemberAccessExprContext) ? (ParserRuleContext)context.children[0] : withExpression.Context, 
                annotations);

            var store = _newFailedResolutionStores.GetOrAdd(declaration.QualifiedModuleName, new ConcurrentFailedResolutionStore());
            store.AddUnresolvedMemberDeclaration(declaration);
        }

        public void AddUnboundDefaultMemberAccess(IdentifierReference defaultMemberAccess)
        {
            var store = _newFailedResolutionStores.GetOrAdd(defaultMemberAccess.QualifiedModuleName, new ConcurrentFailedResolutionStore());
            store.AddUnboundDefaultMemberAccess(defaultMemberAccess);
        }

        public void AddFailedLetCoercionReference(IdentifierReference failedLetCoercion)
        {
            var store = _newFailedResolutionStores.GetOrAdd(failedLetCoercion.QualifiedModuleName, new ConcurrentFailedResolutionStore());
            store.AddFailedLetCoercion(failedLetCoercion);
        }

        public void AddFailedProcedureCoercionReference(IdentifierReference failedProcedureCoercion)
        {
            var store = _newFailedResolutionStores.GetOrAdd(failedProcedureCoercion.QualifiedModuleName, new ConcurrentFailedResolutionStore());
            store.AddFailedProcedureCoercion(failedProcedureCoercion);
        }

        public void AddFailedIndexedDefaultMemberResolution(IdentifierReference failedProcedureCoercion)
        {
            var store = _newFailedResolutionStores.GetOrAdd(failedProcedureCoercion.QualifiedModuleName, new ConcurrentFailedResolutionStore());
            store.AddFailedIndexedDefaultMemberResolution(failedProcedureCoercion);
        }

        public Declaration OnBracketedExpression(string expression, ParserRuleContext context, QualifiedModuleName module)
        {
            var hostApp = FindProject(_hostApp == null ? "VBA" : _hostApp.ApplicationName);
            Debug.Assert(hostApp != null, "Host application project can't be null. Make sure VBA standard library is included if host is unknown.");

            var qualifiedName = module.QualifyMemberName(expression);
            var declarationType = DeclarationType.BracketedExpression;
            var undeclaredScope = (qualifiedName, declarationType);

            if (_newUndeclared.TryGetValue(undeclaredScope, out var undeclared))
            {
                return undeclared.SingleOrDefault();
            }

            var item = new Declaration(qualifiedName, hostApp, hostApp, Tokens.Variant, string.Empty, false, false, Accessibility.Global, DeclarationType.BracketedExpression, context, null, context.GetSelection(), true, null);
            _newUndeclared.TryAdd(undeclaredScope, new ConcurrentBag<Declaration> { item });
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
                      !Declaration.GetModuleParent(p).DeclarationType.HasFlag(DeclarationType.ClassModule)) &&
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
            var match = accessibleMembers.FirstOrDefault(member => !(member.QualifiedModuleName.ComponentName.Equals("Debug") 
                                                                     && (member.IdentifierName.Equals("Assert") 
                                                                         || member.IdentifierName.Equals("Print"))));
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
                        || Declaration.GetModuleParent(p).DeclarationType.HasFlag(DeclarationType.ClassModule)) 
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

        private List<Declaration> FindAllFormControlHandlers()
        {
            var controls = DeclarationsWithType(DeclarationType.Control);
            var handlerNames = BuiltInDeclarations(DeclarationType.Event)
                .SelectMany(e => controls.Select(c => c.IdentifierName + "_" + e.IdentifierName))
                .ToHashSet();
            var handlers = UserDeclarations(DeclarationType.Procedure)
                .Where(procedure => handlerNames.Contains(procedure.IdentifierName));
            return handlers.ToList();
        }

        private HashSet<Declaration> FindAllEventHandlers()
        {
            var handlerNames = BuiltInDeclarations(DeclarationType.Event)
                .SelectMany(e =>
                {
                    var parentModuleSubtypes = ((ClassModuleDeclaration)e.ParentDeclaration).Subtypes.ToList();
                    return parentModuleSubtypes.Any()
                        ? parentModuleSubtypes.Select(v => (v.IdentifierName + "_" + e.IdentifierName).ToLowerInvariant())
                        : new[] { (e.ParentDeclaration.IdentifierName + "_" + e.IdentifierName).ToLowerInvariant() };
                })
                .ToHashSet();

            var handlers = DeclarationsWithType(DeclarationType.Procedure)
                .Where(item =>
                    IsVBAClassSpecificHandler(item) || 
                    IsHostSpecificHandler(item))
                .Concat(
                    UserDeclarations(DeclarationType.Procedure)
                        .Where(item => handlerNames.Any(n => n.Equals(item.IdentifierName, StringComparison.InvariantCultureIgnoreCase)))
                )
                .Concat(_handlersByWithEventsField.Value.AllValues())
                .Concat(FindFormControlEventHandlers())
                .Concat(FindFormEventHandlers())
                .Concat(FindAllDocumentEventHandlers());
            return handlers.ToHashSet();

            // Local functions to help break up the complex logic in finding built-in handlers
            bool IsVBAClassSpecificHandler(Declaration item)
            {
                return item.ParentDeclaration.DeclarationType == DeclarationType.ClassModule && (
                           item.IdentifierName.Equals("Class_Initialize", StringComparison.InvariantCultureIgnoreCase) ||
                           item.IdentifierName.Equals("Class_Terminate", StringComparison.InvariantCultureIgnoreCase));
            }
            
            bool IsHostSpecificHandler(Declaration item)
            {
                return _hostApp?.AutoMacroIdentifiers.Any(i =>
                           i.ComponentTypes.Any(t => t == item.QualifiedModuleName.ComponentType) &&
                           (item.Accessibility != Accessibility.Private || i.MayBePrivate) &&
                           (i.ModuleName == null || i.ModuleName == item.QualifiedModuleName.ComponentName) &&
                           (i.ProcedureName == null || i.ProcedureName == item.IdentifierName)
                       ) ?? false;
            }
        }

        private HashSet<Declaration> FindAllDocumentEventHandlers()
        {
            var documents = DeclarationsWithType(DeclarationType.Document).OfType<DocumentModuleDeclaration>();
            var documentTypes = documents.SelectMany(doc => doc.Supertypes).ToHashSet();
            var events = BuiltInDeclarations(DeclarationType.Event).OfType<EventDeclaration>().Where(e => documentTypes.Contains(e.ParentDeclaration));
            var handlerNames = events.Select(e => (Event:e, HandlerName:$"{e.QualifiedModuleName.ComponentName}_{e.IdentifierName}".ToLowerInvariant())).ToHashSet();

            var procedures = UserDeclarations(DeclarationType.Procedure)
                .Where(procedure => procedure.ParentDeclaration is DocumentModuleDeclaration);

            var candidates = procedures.Where(procedure => handlerNames.Select(e => e.HandlerName).Contains(procedure.IdentifierName.ToLowerInvariant()))
                .Select(c => (CandidateHandler: c as SubroutineDeclaration, Event: handlerNames.Single(h => h.HandlerName.EndsWith(c.IdentifierName)).Event));

            var handlers = candidates.Where(candidate => candidate.CandidateHandler.Parameters.Count == candidate.Event.Parameters.Count
                && Enumerable.SequenceEqual(candidate.CandidateHandler.Parameters.Select(p => p.FullAsTypeName), 
                                            candidate.Event.Parameters.Select(p => p.FullAsTypeName)));
            
            return handlers.Select(h => h.CandidateHandler).Cast<Declaration>().ToHashSet();
        }

        private HashSet<Declaration> FindAllFormEventHandlers()
        {
            var forms = DeclarationsWithType(DeclarationType.ClassModule)
                .Where(declaration => declaration.QualifiedModuleName.ComponentType == ComponentType.UserForm);
            var formScopes = forms
                .Select(form => form.Scope)
                .ToHashSet();
            var events = BuiltInDeclarations(DeclarationType.Event)
                .Where(item => item.ParentScope == "FM20.DLL;MSForms.FormEvents");
            var handlerNames = events
                .Select(item => ("UserForm_" + item.IdentifierName).ToLowerInvariant())
                .ToHashSet();
            var handlers = UserDeclarations(DeclarationType.Procedure)
                .Where(procedure => handlerNames.Contains(procedure.IdentifierName.ToLowerInvariant())
                                    && formScopes.Contains(procedure.ParentScope));
            return handlers.ToHashSet();
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

            var identifierMatches = MatchName(newName).Where(match => match.ProjectId == renameTarget.ProjectId).ToList();

            if (!identifierMatches.Any())
            {
                return Enumerable.Empty<Declaration>();
            }

            if (IsEnumOrUDTMemberDeclaration(renameTarget)) 
            {
                return identifierMatches.Where(idm =>
                    IsEnumOrUDTMemberDeclaration(idm) && idm.ParentDeclaration == renameTarget.ParentDeclaration);
            }

            identifierMatches = identifierMatches.Where(nc => !IsEnumOrUDTMemberDeclaration(nc)).ToList();
            var referenceConflicts = identifierMatches.Where(idm =>
                renameTarget.References
                    .Any(renameTargetRef => 
                        renameTargetRef.ParentScoping == idm.ParentDeclaration
                        
                        || !renameTarget.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule)
                            && idm == renameTargetRef.ParentScoping
                            && !UsesScopeResolution(renameTargetRef.Context.Parent)

                        || idm.References
                            .Any(idmRef => idmRef.ParentScoping == renameTargetRef.ParentScoping
                                && renameTargetRef.QualifiedModuleName != renameTarget.QualifiedModuleName
                                && !UsesScopeResolution(renameTargetRef.Context.Parent))

                        || idm.References
                            .Any(idmRef => idmRef.ParentScoping == renameTargetRef.ParentScoping
                                && renameTargetRef.QualifiedModuleName == renameTarget.QualifiedModuleName
                                && !UsesScopeResolution(idmRef.Context.Parent)))

                || idm.DeclarationType.HasFlag(DeclarationType.Variable)
                    && idm.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module)
                    && renameTarget.References.Any(renameTargetRef => renameTargetRef.QualifiedModuleName == idm.ParentDeclaration.QualifiedModuleName))
                    .ToList();

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
                    && IsConflictingMember(renameTarget, renameTargetModule, idm))
                    .ToList();

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

        public bool IsReferenceUsedInProject(ProjectDeclaration project, ReferenceInfo reference, bool checkForward = false)
        {
            if (project == null || string.IsNullOrEmpty(reference.FullPath))
            {
                return false;
            }

            var referenceProject = GetProjectDeclarationForReference(reference);

            if (referenceProject == null ||         // Can't locate the project for the reference - assume it is used to avoid false negatives.
                IdentifierReferences().Any(item =>
                item.Key.ProjectId == project.ProjectId && item.Value.Any(usage =>
                    ReferenceEquals(Declaration.GetProjectParent(usage.Declaration), project))))
            {
                return true;
            }

            if (!checkForward)
            {
                return false;
            }

            // OK, no direct references - check indirect references in built-in libraries (i.e. Excel forward references Office)
            return !referenceProject.IsUserDefined && AllBuiltInDeclarations.Any(declaration =>
                       declaration.AsTypeDeclaration != null &&
                       declaration.AsTypeDeclaration.QualifiedModuleName.ProjectId.Equals(referenceProject.ProjectId));
        }

        public List<IdentifierReference> FindAllReferenceUsesInProject(ProjectDeclaration project, ReferenceInfo reference, 
            out ProjectDeclaration referenceProject)
        {
            var output = new List<IdentifierReference>();
            if (project == null || string.IsNullOrEmpty(reference.FullPath))
            {
                referenceProject = null;
                return output;
            }

            referenceProject = GetProjectDeclarationForReference(reference);
            if (referenceProject == null)
            {
                Logger.Warn($"Could not get the project declaration for reference '{reference.Name}'.");
                return output;
            }
            if (!_referencesByProjectId.TryGetValue(referenceProject.ProjectId, out var directReferences))
            {
                return output;
            }
            output.AddRange(directReferences);

            var projectId = referenceProject.ProjectId;

            output.AddRange(_identifierReferences.Where(identifier =>
                identifier?.Declaration?.AsTypeDeclaration != null &&
                identifier.Declaration.AsTypeDeclaration.QualifiedModuleName.ProjectId.Equals(projectId)));

            return output;
        }

        private ProjectDeclaration GetProjectDeclarationForReference(ReferenceInfo reference)
        {
            return reference.Guid.Equals(Guid.Empty)
                ? UserDeclarations(DeclarationType.Project).OfType<ProjectDeclaration>().FirstOrDefault(proj =>
                    proj.QualifiedModuleName.ProjectPath.Equals(reference.FullPath, StringComparison.InvariantCultureIgnoreCase))
                : BuiltInDeclarations(DeclarationType.Project).OfType<ProjectDeclaration>().FirstOrDefault(proj =>
                    proj.Guid.Equals(reference.Guid) && proj.MajorVersion == reference.Major &&
                    proj.MinorVersion == reference.Minor);
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
        /// Gets all identifier references for an IdentifierContext.
        /// </summary>
        public IEnumerable<IdentifierReference> IdentifierReferences(VBAParser.IdentifierContext identifierContext, QualifiedModuleName module)
        {
            var qualifiedSelection = new QualifiedSelection(module, identifierContext.GetSelection());
            return IdentifierReferences(qualifiedSelection)
                .Where(identifierReference => identifierReference.IdentifierName.Equals(identifierContext.GetText()));
        }

        /// <summary>
        /// Gets all identifier references for an UnrestrictedIdentifierContext.
        /// </summary>
        public IEnumerable<IdentifierReference> IdentifierReferences(VBAParser.UnrestrictedIdentifierContext identifierContext, QualifiedModuleName module)
        {
            var qualifiedSelection = new QualifiedSelection(module, identifierContext.GetSelection());
            return IdentifierReferences(qualifiedSelection)
                .Where(identifierReference => identifierReference.IdentifierName.Equals(identifierContext.GetText()));
        }

        /// <summary>
        /// Gets all identifier references with the specified selection.
        /// </summary>
        public IEnumerable<IdentifierReference> IdentifierReferences(QualifiedSelection selection)
        {
            return _referencesBySelection.TryGetValue(selection, out var value)
                ? value.OrderBy(reference => reference.DefaultMemberRecursionDepth)
                : Enumerable.Empty<IdentifierReference>();
        }

        /// <summary>
        /// Gets all identifier references within a qualified selection, ordered by selection (start position, then length)
        /// </summary>
        public IEnumerable<IdentifierReference> ContainedIdentifierReferences(QualifiedSelection qualifiedSelection)
        {
            return IdentifierReferences(qualifiedSelection.QualifiedName)
                .Where(reference => qualifiedSelection.Selection.Contains(reference.Selection))
                .OrderBy(reference => reference.Selection)
                .ThenBy(reference => reference.DefaultMemberRecursionDepth);
        }

        /// <summary>
        /// Gets all identifier references containing a qualified selection, ordered by selection (start position, then length).
        /// Default member accesses with identical selections are ordered by call order.
        /// </summary>
        public IEnumerable<IdentifierReference> ContainingIdentifierReferences(QualifiedSelection qualifiedSelection)
        {
            return IdentifierReferences(qualifiedSelection.QualifiedName)
                .Where(reference => reference.Selection.Contains(qualifiedSelection.Selection))
                .OrderBy(reference => reference.Selection)
                .ThenBy(reference => reference.DefaultMemberRecursionDepth);
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

        /// <summary>
        /// Gets the unbound default member calls in a module.
        /// </summary>
        public IReadOnlyCollection<IdentifierReference> UnboundDefaultMemberAccesses(QualifiedModuleName module)
        {
            return _failedResolutionStores.TryGetValue(module, out var store)
                ? store.UnboundDefaultMemberAccesses
                : new HashSet<IdentifierReference>();
        }

        /// <summary>
        /// Gets all unbound default member calls.
        /// </summary>
        public IEnumerable<IdentifierReference> AllUnboundDefaultMemberAccesses()
        {
            return _failedResolutionStores.Values
                .SelectMany(store => store.UnboundDefaultMemberAccesses);
        }

        /// <summary>
        /// Gets the failed Let coercions in a module.
        /// </summary>
        public IReadOnlyCollection<IdentifierReference> FailedLetCoercions(QualifiedModuleName module)
        {
            return _failedResolutionStores.TryGetValue(module, out var store)
                ? store.FailedLetCoercions
                : new HashSet<IdentifierReference>();
        }

        /// <summary>
        /// Gets all failed Let coercions.
        /// </summary>
        public IEnumerable<IdentifierReference> FailedLetCoercions()
        {
            return _failedResolutionStores.Values
                .SelectMany(store => store.FailedLetCoercions);
        }

        /// <summary>
        /// Gets the failed procedure coercions in a module.
        /// </summary>
        public IReadOnlyCollection<IdentifierReference> FailedProcedureCoercions(QualifiedModuleName module)
        {
            return _failedResolutionStores.TryGetValue(module, out var store)
                ? store.FailedProcedureCoercions
                : new HashSet<IdentifierReference>();
        }

        /// <summary>
        /// Gets all failed procedure coercions.
        /// </summary>
        public IEnumerable<IdentifierReference> FailedProcedureCoercions()
        {
            return _failedResolutionStores.Values
                .SelectMany(store => store.FailedProcedureCoercions);
        }

        /// <summary>
        /// Gets the unresolved member call declarations to a members in a module.
        /// </summary>
        public IReadOnlyCollection<UnboundMemberDeclaration> UnresolvedMemberDeclarations(QualifiedModuleName module)
        {
            return _failedResolutionStores.TryGetValue(module, out var store)
                ? store.UnresolvedMemberDeclarations
                : new HashSet<UnboundMemberDeclaration>();
        }

        /// <summary>
        /// Gets all unresolved member call declarations.
        /// </summary>
        public IEnumerable<UnboundMemberDeclaration> UnresolvedMemberDeclarations()
        {
            return _failedResolutionStores.Values
                .SelectMany(store => store.UnresolvedMemberDeclarations);
        }

        /// <summary>
        /// Gets the failed indexed default member accesses in a module.
        /// </summary>
        public IReadOnlyCollection<IdentifierReference> FailedIndexedDefaultMemberAccesses(QualifiedModuleName module)
        {
            return _failedResolutionStores.TryGetValue(module, out var store)
                ? store.FailedIndexedDefaultMemberResolutions
                : new HashSet<IdentifierReference>();
        }

        /// <summary>
        /// Gets all failed indexed default member accesses.
        /// </summary>
        public IEnumerable<IdentifierReference> FailedIndexedDefaultMemberAccesses()
        {
            return _failedResolutionStores.Values
                .SelectMany(store => store.FailedIndexedDefaultMemberResolutions);
        }
    }
}
