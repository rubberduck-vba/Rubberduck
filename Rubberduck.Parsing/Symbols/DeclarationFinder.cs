using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.VBEditor;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols
{
    internal static class DictionaryExtensions
    {
        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this ConcurrentDictionary<TKey, ConcurrentBag<TValue>> source)
        {
            return source.SelectMany(item => item.Value).ToList();
        }

        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this IDictionary<TKey, IList<TValue>> source)
        {
            return source.SelectMany(item => item.Value).ToList();
        }

        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this IDictionary<TKey, List<TValue>> source)
        {
            return source.SelectMany(item => item.Value);
        }

        public static ConcurrentDictionary<TKey, ConcurrentBag<TValue>> ToConcurrentDictionary<TKey, TValue>(this IEnumerable<IGrouping<TKey, TValue>> source)
        {
            return new ConcurrentDictionary<TKey, ConcurrentBag<TValue>>(source.Select(x => new KeyValuePair<TKey, ConcurrentBag<TValue>>(x.Key, new ConcurrentBag<TValue>(x))));
        }

        public static Dictionary<TKey, List<TValue>> ToDictionary<TKey, TValue>(this IEnumerable<IGrouping<TKey, TValue>> source)
        {
            return source.ToDictionary(group => group.Key, group => group.ToList());
        }
    }

    public class DeclarationFinder
    {
        private static readonly SquareBracketedNameComparer NameComparer = new SquareBracketedNameComparer();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IHostApplication _hostApp;
        private readonly AnnotationService _annotationService;
        private IDictionary<string, List<Declaration>> _declarationsByName;
        private IDictionary<QualifiedModuleName, List<Declaration>> _declarations;
        private ConcurrentDictionary<QualifiedMemberName, ConcurrentBag<Declaration>> _newUndeclared;
        private ConcurrentBag<UnboundMemberDeclaration> _newUnresolved;
        private List<UnboundMemberDeclaration> _unresolved;
        private IDictionary<QualifiedModuleName, List<IAnnotation>> _annotations;
        private IDictionary<Declaration, List<Declaration>> _parametersByParent;
        private IDictionary<DeclarationType, List<Declaration>> _userDeclarationsByType;
        private IDictionary<QualifiedSelection, List<Declaration>> _declarationsBySelection;
        private IDictionary<QualifiedSelection, List<IdentifierReference>> _referencesBySelection;

        private Lazy<IDictionary<DeclarationType, List<Declaration>>> _builtInDeclarationsByType;
        private Lazy<IDictionary<Declaration, List<Declaration>>> _handlersByWithEventsField;
        private Lazy<IDictionary<VBAParser.ImplementsStmtContext, List<Declaration>>> _membersByImplementsContext;
        private Lazy<IDictionary<Declaration, List<Declaration>>> _interfaceMembers;
        private Lazy<List<Declaration>> _nonBaseAsType;
        private Lazy<List<Declaration>> _eventHandlers;
        private Lazy<List<Declaration>> _projects;
        private Lazy<List<Declaration>> _classes;
        
        private readonly object threadLock = new object();

        private static QualifiedSelection GetGroupingKey(Declaration declaration)
        {
            // we want the procedures' whole body, not just their identifier:
            return declaration.DeclarationType.HasFlag(DeclarationType.Member)
                ? new QualifiedSelection(
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.Context.GetSelection())
                : declaration.QualifiedSelection;
        }

        public DeclarationFinder(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, IHostApplication hostApp = null)
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

        private List<Action> CollectionConstructionActions(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations)
        {
            var actions = new List<Action>();

            actions.Add(() => 
                _unresolved = unresolvedMemberDeclarations
                    .ToList()
                );
            actions.Add(() =>
                _annotations = annotations
                    .GroupBy(node => node.QualifiedSelection.QualifiedName)
                    .ToDictionary()
                );
            actions.Add(() => 
                _declarations = declarations
                    .GroupBy(item => item.QualifiedName.QualifiedModuleName)
                    .ToDictionary()
                );
            actions.Add(() => 
                _declarationsByName = declarations
                    .GroupBy(declaration => declaration.IdentifierName.ToLowerInvariant())
                    .ToDictionary()
                );
            actions.Add(() =>
                _declarationsBySelection = declarations
                    .Where(declaration => declaration.IsUserDefined)
                    .GroupBy(GetGroupingKey)
                    .ToDictionary()
                );
            actions.Add(() => 
                _referencesBySelection = declarations
                    .SelectMany(declaration => declaration.References)
                    .GroupBy(reference => new QualifiedSelection(reference.QualifiedModuleName, reference.Selection))
                    .ToDictionary()
                );
            actions.Add(() =>
                _parametersByParent = declarations
                    .Where(declaration => declaration.DeclarationType == DeclarationType.Parameter)
                    .GroupBy(declaration => declaration.ParentDeclaration)
                    .ToDictionary()
                );
            actions.Add(() =>
                _userDeclarationsByType = declarations
                    .Where(declaration => declaration.IsUserDefined)
                    .GroupBy(declaration => declaration.DeclarationType)
                    .ToDictionary()
                );

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

            _eventHandlers = new Lazy<List<Declaration>>(() => FindAllEventHandlers(), true);
            _projects = new Lazy<List<Declaration>>(() => DeclarationsWithType(DeclarationType.Project).ToList(), true);
            _classes = new Lazy<List<Declaration>>(() => DeclarationsWithType(DeclarationType.ClassModule).ToList(), true);
            _handlersByWithEventsField = new Lazy<IDictionary<Declaration, List<Declaration>>>(() => FindAllHandlersByWithEventField(), true);
            _interfaceMembers = new Lazy<IDictionary<Declaration, List<Declaration>>>(() => FindAllIinterfaceMembersByModule(), true);
            _membersByImplementsContext = new Lazy<IDictionary<VBAParser.ImplementsStmtContext, List<Declaration>>>(() =>
                FindAllImplementingMembersByImplementsContext(),
                true);
        }

        private IDictionary<VBAParser.ImplementsStmtContext, List<Declaration>> FindAllImplementingMembersByImplementsContext()
        {
            var implementsInstructions = UserDeclarations(DeclarationType.ClassModule)
                .SelectMany(cls => cls.References
                    .Where(reference => reference.Context is VBAParser.ImplementsStmtContext 
                        || (reference.Context).IsDescendentOf<VBAParser.ImplementsStmtContext>())
                    .Select(reference =>
                        new
                        {
                            IdentifierReference = reference,
                            Context = reference.Context is VBAParser.ImplementsStmtContext ? 
                                (VBAParser.ImplementsStmtContext)reference.Context 
                                    : reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>()
                        }
                    )
                ).ToList();

            var implementingNames = implementsInstructions.SelectMany(item =>
                    Members(item.IdentifierReference.Declaration.QualifiedName.QualifiedModuleName)
                        .Where(member => member.DeclarationType.HasFlag(DeclarationType.Member))
                        .Select(member => item.IdentifierReference.Declaration.IdentifierName + "_" + member.IdentifierName)
                        )
                    .ToHashSet();

            var implementableMembers = implementsInstructions.Select(item =>
                new
                {
                    item.Context,
                    Members = Members(item.IdentifierReference.QualifiedModuleName)
                        .Where(implementingTypeMember => implementingNames.Contains(implementingTypeMember.IdentifierName))
                        .ToList()
                });

            return implementableMembers.ToDictionary(item => item.Context, item => item.Members);
        }

        private IDictionary<Declaration, List<Declaration>> FindAllIinterfaceMembersByModule()
        {
            var implementsInstructions = UserDeclarations(DeclarationType.ClassModule)
                .SelectMany(cls => cls.References
                    .Where(reference => reference.Context is VBAParser.ImplementsStmtContext || reference.Context.IsDescendentOf<VBAParser.ImplementsStmtContext>())
                    .Select(reference =>
                        new
                        {
                            IdentifierReference = reference,
                            Context = reference.Context is VBAParser.ImplementsStmtContext ?
                                (VBAParser.ImplementsStmtContext)reference.Context
                                : reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>()
                        }
                    )
                );

            var interfaceModules = implementsInstructions.Select(item => item.IdentifierReference.Declaration).Distinct();

            var interfaceMembers = interfaceModules.Select(item =>
                new
                {
                    InterfaceModule = item,
                    InterfaceMembers = Members(item.QualifiedName.QualifiedModuleName)
                        .Where(member => member.DeclarationType.HasFlag(DeclarationType.Member))
                });

            return interfaceMembers.ToDictionary(
                        item => item.InterfaceModule,
                        item => item.InterfaceMembers.ToList()
                     );
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
            if (!qualifiedSelection.HasValue || qualifiedSelection.Value.Equals(default(QualifiedSelection)))
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
            List<Declaration> members;
            return _declarations.TryGetValue(module, out members)
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
            List<Declaration> result;
            return _userDeclarationsByType.TryGetValue(type, out result)
                ? result
                : _userDeclarationsByType
                    .Where(item => item.Key.HasFlag(type))
                    .SelectMany(item => item.Value);
        }

        public IEnumerable<Declaration> AllUserDeclarations => _userDeclarationsByType.AllValues();

        public IEnumerable<Declaration> BuiltInDeclarations(DeclarationType type)
        {
            List<Declaration> result;
            return _builtInDeclarationsByType.Value.TryGetValue(type, out result)
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
            List<Declaration> result;
            return _handlersByWithEventsField.Value.TryGetValue(field, out result) 
                ? result 
                : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<Declaration> FindInterfaceMembersForImplementsContext(VBAParser.ImplementsStmtContext context)
        {
            List<Declaration> result;
            return _membersByImplementsContext.Value.TryGetValue(context, out result)
                ? result
                : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<Declaration> FindAllInterfaceMembers()
        {
            return _interfaceMembers.Value.SelectMany(item => item.Value);
        }

        public IEnumerable<Declaration> FindAllInterfaceImplementingMembers()
        {
            return _membersByImplementsContext.Value.AllValues();
        }

        public Declaration FindParameter(Declaration procedure, string parameterName)
        {
            List<Declaration> parameters;
            return _parametersByParent.TryGetValue(procedure, out parameters) 
                ? parameters.SingleOrDefault(parameter => parameter.IdentifierName == parameterName) 
                : null;
        }

        public IEnumerable<Declaration> FindMemberMatches(Declaration parent, string memberName)
        {
            List<Declaration> children;
            return _declarations.TryGetValue(parent.QualifiedName.QualifiedModuleName, out children)
                ? children.Where(item => item.DeclarationType.HasFlag(DeclarationType.Member)
                                             && item.IdentifierName == memberName)
                : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<IAnnotation> FindAnnotations(QualifiedModuleName module)
        {
            List<IAnnotation> result;
            return _annotations.TryGetValue(module, out result) 
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
            List<Declaration> result;
            return _declarationsByName.TryGetValue(normalizedName, out result) 
                ? result 
                : Enumerable.Empty<Declaration>();
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

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => referencedProject.Equals(Declaration.GetProjectParent(p)) && p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => p.DeclarationType == DeclarationType.ClassModule && ((ClassModuleDeclaration)p).HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName,
                p => referencedProject.Equals(Declaration.GetProjectParent(p))
                    && p.DeclarationType == DeclarationType.ClassModule 
                    && ((ClassModuleDeclaration)p).HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberWithParent(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration parent, string memberName, DeclarationType memberType)
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

        public Declaration FindMemberEnclosingProcedure(Declaration enclosingProcedure, string memberName, DeclarationType memberType, ParserRuleContext onSiteContext = null)
        {
            if (memberType == DeclarationType.Variable && NameComparer.Equals(enclosingProcedure.IdentifierName, memberName))
            {
                return enclosingProcedure;
            }
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && enclosingProcedure.Equals(m.ParentDeclaration));
            return memberMatches.FirstOrDefault();
        }

        public Declaration OnUndeclaredVariable(Declaration enclosingProcedure, string identifierName, ParserRuleContext context)
        {
            var annotations = _annotationService.FindAnnotations(enclosingProcedure.QualifiedName.QualifiedModuleName, context.Start.Line);
            var undeclaredLocal =
                new Declaration(
                    new QualifiedMemberName(enclosingProcedure.QualifiedName.QualifiedModuleName, identifierName),
                    enclosingProcedure, enclosingProcedure, "Variant", string.Empty, false, false,
                    Accessibility.Implicit, DeclarationType.Variable, context, context.GetSelection(), false, null,
                    true, annotations, null, true);

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

            ConcurrentBag<Declaration> undeclared;
            if (_newUndeclared.TryGetValue(qualifiedName, out undeclared))
            {
                return undeclared.SingleOrDefault();
            }

            var item = new Declaration(qualifiedName, hostApp, hostApp, Tokens.Variant, string.Empty, false, false, Accessibility.Global, DeclarationType.BracketedExpression, context, context.GetSelection(), true, null);
            _newUndeclared.TryAdd(qualifiedName, new ConcurrentBag<Declaration> { item });
            return item;
        }

        public Declaration FindMemberEnclosedProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
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

        public Declaration FindMemberEnclosedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, string memberName, DeclarationType memberType)
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
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => (!isInstanceSensitive || Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType != DeclarationType.ClassModule) && p.DeclarationType.HasFlag(memberType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, DeclarationType moduleType, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && (Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType == moduleType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInGlobalClassModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
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

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, string memberName, DeclarationType memberType)
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

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration referencedProject, string memberName, DeclarationType memberType)
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
                    var parentModuleSubtypes = ((ClassModuleDeclaration)e.ParentDeclaration).Subtypes;
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


        public IEnumerable<Declaration> GetAccessibleDeclarations(Declaration target)
        {
            if (target == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            return _declarations.AllValues()
                .Where(callee => AccessibilityCheck.IsAccessible(
                    Declaration.GetProjectParent(target),
                    Declaration.GetModuleParent(target), 
                    target.ParentDeclaration, 
                    callee));
        }

        public IEnumerable<Declaration> GetAccessibleUserDeclarations(Declaration target)
        {
            if (target == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            return _userDeclarationsByType.AllValues()
                .Where(callee => AccessibilityCheck.IsAccessible(
                    Declaration.GetProjectParent(target),
                    Declaration.GetModuleParent(target),
                    target.ParentDeclaration,
                    callee));
        }

        public IEnumerable<Declaration> GetDeclarationsWithIdentifiersToAvoid(Declaration target)
        {
            if (target == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            List<Declaration> declarationsToAvoid = GetNameCollisionDeclarations(target).ToList();

            declarationsToAvoid.AddRange(GetNameCollisionDeclarations(target.References));

            return declarationsToAvoid.Distinct();
        }

        private IEnumerable<Declaration> GetNameCollisionDeclarations(Declaration declaration)
        {
            if (declaration == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            //Filter accessible declarations to those that would result in name collisions or hiding
            var declarationsToAvoid = GetAccessibleUserDeclarations(declaration).Where(candidate =>
                                        (IsAccessibleInOtherProcedureModule(candidate,declaration)
                                        || candidate.DeclarationType == DeclarationType.Project
                                        || candidate.DeclarationType.HasFlag(DeclarationType.Module)
                                        || IsDeclarationInSameProcedureScope(candidate, declaration)
                                        )).ToHashSet();

            //Add local variables when the target is a method or property
            if(IsSubroutineOrProperty(declaration))
            {
                var localVariableDeclarations = _declarations.AllValues()
                    .Where(dec => declaration == dec.ParentDeclaration);
                declarationsToAvoid.UnionWith(localVariableDeclarations);
            }

            return declarationsToAvoid;
        }

        private IEnumerable<Declaration> GetNameCollisionDeclarations(IEnumerable<IdentifierReference> references)
        {
            var declarationsToAvoid = new HashSet<Declaration>();
            foreach (var reference in references)
            {
                if (!UsesScopeResolution(reference.Context.Parent))
                {
                    declarationsToAvoid.UnionWith(GetNameCollisionDeclarations(reference.ParentNonScoping));
                }
            }
            return declarationsToAvoid;
        }

        private bool IsAccessibleInOtherProcedureModule(Declaration candidate, Declaration declaration)
        {
            return IsInProceduralModule(declaration)
                       && IsInProceduralModule(candidate)
                       && candidate.Accessibility != Accessibility.Private;
        }

        private bool UsesScopeResolution(RuleContext ruleContext)
        {
            return (ruleContext is VBAParser.WithMemberAccessExprContext)
                || (ruleContext is VBAParser.MemberAccessExprContext);
        }

        private bool IsInProceduralModule(Declaration candidateDeclaration)
        {
            var candidateModuleDeclaration = Declaration.GetModuleParent(candidateDeclaration);

            return candidateModuleDeclaration?.DeclarationType == DeclarationType.ProceduralModule;
        }

        private bool IsDeclarationInSameProcedureScope(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return candidateDeclaration.ParentScope == scopingDeclaration.ParentScope;
        }

        private static bool IsSubroutineOrProperty(Declaration declaration)
        {
            return declaration.DeclarationType.HasFlag(DeclarationType.Property)
                || declaration.DeclarationType == DeclarationType.Function
                || declaration.DeclarationType == DeclarationType.Procedure;
        }
    }
}
