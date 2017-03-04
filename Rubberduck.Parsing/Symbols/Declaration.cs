using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines a declared identifier.
    /// </summary>
    [DebuggerDisplay("({DeclarationType}) {Accessibility} {IdentifierName} As {AsTypeName} | {Selection}")]
    public class Declaration : IEquatable<Declaration>
    {
        public Declaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            Declaration parentScope,
            string asTypeName,
            string typeHint,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            DeclarationType declarationType,
            ParserRuleContext context,
            Selection selection,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            bool isBuiltIn = true,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null,
            bool undeclared = false)
            : this(
                qualifiedName,
                parentDeclaration,
                parentScope == null ? null : parentScope.Scope,
                asTypeName,
                typeHint,
                isSelfAssigned,
                isWithEvents,
                accessibility,
                declarationType,
                context,
                selection,
                isArray,
                asTypeContext,
                isBuiltIn,
                annotations,
                attributes)
        {
            _parentScopeDeclaration = parentScope;
            _undeclared = undeclared;
        }

        public Declaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            string parentScope,
            string asTypeName,
            string typeHint,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            DeclarationType declarationType,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            bool isBuiltIn = true,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null)
            : this(
                  qualifiedName,
                  parentDeclaration,
                  parentScope,
                  asTypeName,
                  typeHint,
                  isSelfAssigned,
                  isWithEvents,
                  accessibility,
                  declarationType,
                  null,
                  Selection.Home,
                  isArray,
                  asTypeContext,
                  isBuiltIn,
                  annotations,
                  attributes)
        { }

        public Declaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            string parentScope,
            string asTypeName,
            string typeHint,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            DeclarationType declarationType,
            ParserRuleContext context,
            Selection selection,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            bool isBuiltIn = false,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null)
        {
            _qualifiedName = qualifiedName;            
            _parentDeclaration = parentDeclaration;
            _parentScopeDeclaration = _parentDeclaration;
            _parentScope = parentScope ?? string.Empty;
            _identifierName = qualifiedName.MemberName;
            _asTypeName = asTypeName;
            _isSelfAssigned = isSelfAssigned;
            _isWithEvents = isWithEvents;
            _accessibility = accessibility;
            _declarationType = declarationType;
            _selection = selection;
            Context = context;
            _isBuiltIn = isBuiltIn;
            _annotations = annotations;
            _attributes = attributes ?? new Attributes();

            _projectId = _qualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = GetProjectParent(parentDeclaration);
            if (projectDeclaration != null)
            {
                _projectName = projectDeclaration.IdentifierName;
            }
            else if (_declarationType == DeclarationType.Project)
            {
                _projectName = _identifierName;
            }

            _customFolder = FolderFromAnnotations();
            _isArray = isArray;
            _asTypeContext = asTypeContext;
            _typeHint = typeHint;
        }

        public Declaration(ComEnumeration enumeration, Declaration parent, QualifiedModuleName module) : this(
            module.QualifyMemberName(enumeration.Name),
            parent,
            parent,
            "Long",
            //Match the VBA default type declaration.  Technically these *can* be a LongLong on 64 bit systems, but would likely crash the VBE... 
            null,
            false,
            false,
            Accessibility.Global,
            DeclarationType.Enumeration,
            null,
            Selection.Home,
            false,
            null,
            true,
            null,
            new Attributes()) { }

        public Declaration(ComStruct structure, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName(structure.Name),
                parent,
                parent,
                structure.Name,
                null,
                false,
                false,
                Accessibility.Global,
                DeclarationType.UserDefinedType,
                null,
                Selection.Home,
                false,
                null,
                true,
                null,
                new Attributes()) { }

        public Declaration(ComEnumerationMember member, Declaration parent, QualifiedModuleName module) : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                parent.IdentifierName,
                null,
                false,
                false,
                Accessibility.Global,
                DeclarationType.EnumerationMember,
                null,
                Selection.Home,
                false,
                null) { }

        public Declaration(ComField field, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName(field.Name),
                parent,
                parent,
                field.ValueType,
                null,
                false,
                false,
                Accessibility.Global,
                field.Type,
                null,
                Selection.Home,
                false,
                null) { }

        private string FolderFromAnnotations()
            {
                var @namespace = Annotations.FirstOrDefault(annotation => annotation.AnnotationType == AnnotationType.Folder);
                string result;
                if (@namespace == null)
                {
                    result = string.IsNullOrEmpty(_qualifiedName.QualifiedModuleName.ProjectName)
                        ? _projectId
                        : _qualifiedName.QualifiedModuleName.ProjectName;
                }
                else
                {
                    var value = ((FolderAnnotation)@namespace).FolderName;
                    result = value;
                }
                return result;
            }


        public static Declaration GetModuleParent(Declaration declaration)
        {
            if (declaration == null)
            {
                return null;
            }
            if (declaration.DeclarationType == DeclarationType.ClassModule || declaration.DeclarationType == DeclarationType.ProceduralModule)
            {
                return declaration;
            }
            return GetModuleParent(declaration.ParentDeclaration);
        }

        public static Declaration GetProjectParent(Declaration declaration)
        {
            if (declaration == null)
            {
                return null;
            }
            if (declaration.DeclarationType == DeclarationType.Project)
            {
                return declaration;
            }
            return GetProjectParent(declaration.ParentDeclaration);
        }

        private readonly bool _isArray;
        public bool IsArray { get { return _isArray; } }
        private readonly VBAParser.AsTypeClauseContext _asTypeContext;
        public VBAParser.AsTypeClauseContext AsTypeContext { get { return _asTypeContext; } }
        private readonly string _typeHint;
        public string TypeHint { get { return _typeHint; } }
        public bool HasTypeHint { get { return !string.IsNullOrWhiteSpace(_typeHint); } }

        public bool IsTypeSpecified
        {
            get
            {
                return HasTypeHint || _asTypeContext != null;
            }
        }

        private readonly bool _isBuiltIn;
        public bool IsBuiltIn { get { return _isBuiltIn; } }

        private readonly Declaration _parentDeclaration;
        public Declaration ParentDeclaration { get { return _parentDeclaration; } }

        private readonly QualifiedMemberName _qualifiedName;
        public QualifiedMemberName QualifiedName { get { return _qualifiedName; } }

        public ParserRuleContext Context { get; set; }

        private ConcurrentBag<IdentifierReference> _references = new ConcurrentBag<IdentifierReference>();
        public IEnumerable<IdentifierReference> References
        {
            get
            {
                return _references;
            }
            set
            {
                _references = new ConcurrentBag<IdentifierReference>(value);
            }
        }

        private readonly IEnumerable<IAnnotation> _annotations;
        public IEnumerable<IAnnotation> Annotations { get { return _annotations ?? new List<IAnnotation>(); } }

        private readonly Attributes _attributes;
        public IReadOnlyDictionary<string, IEnumerable<string>> Attributes { get { return _attributes; } }

        /// <summary>
        /// Gets an attribute value that contains the docstring for a member.
        /// </summary>
        public string DescriptionString
        {
            get
            {
                IEnumerable<string> value;
                if (_attributes.TryGetValue("VB_Description", out value))
                {
                    return value.Single();
                }

                return string.Empty;
            }
        }

        /// <summary>
        /// Gets an attribute value indicating whether a member is an enumerator provider.
        /// Types with such a member support For Each iteration.
        /// </summary>
        public bool IsEnumeratorMember
        {
            get
            {
                IEnumerable<string> value;
                if (_attributes.TryGetValue("VB_UserMemId", out value))
                {
                    return value.Single() == "-4";
                }

                return false;
            }
        }

        public void AddReference(
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            ParserRuleContext callSiteContext,
            string identifier,
            Declaration callee,
            Selection selection,
            IEnumerable<IAnnotation> annotations,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false)
        {
            _references.Add(
                new IdentifierReference(
                    module,
                    scope,
                    parent,
                    identifier,
                    selection,
                    callSiteContext,
                    callee,
                    isAssignmentTarget,
                    hasExplicitLetStatement,
                    annotations));
        }

        private readonly Selection _selection;
        /// <summary>
        /// Gets a <c>Selection</c> representing the position of the declaration in the code module.
        /// </summary>
        /// <remarks>
        /// Returns <c>default(Selection)</c> for module identifiers.
        /// </remarks>
        public Selection Selection { get { return _selection; } }

        public QualifiedSelection QualifiedSelection { get { return new QualifiedSelection(_qualifiedName.QualifiedModuleName, _selection); } }

        /// <summary>
        /// Gets a reference to the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public virtual IVBProject Project { get { return _parentDeclaration.Project; } }

        private readonly string _projectId;
        /// <summary>
        /// Gets a unique identifier for the VBProject the declaration is made in.
        /// </summary>
        public string ProjectId { get { return _projectId; } }

        private readonly string _projectName;
        public string ProjectName
        {
            get { return _projectName; }
        }

        /// <summary>
        /// WARNING: This property has side effects. It changes the ActiveVBProject, which causes a flicker in the VBE.
        /// This should only be called if it is *absolutely* necessary.
        /// </summary>
        public virtual string ProjectDisplayName { get { return _parentDeclaration.ProjectDisplayName; } }

        public object[] ToArray()
        {
            return new object[] { ProjectName, CustomFolder, ComponentName, DeclarationType.ToString(), Scope, IdentifierName, AsTypeName };
        }


        /// <summary>
        /// Gets the name of the VBComponent the declaration is made in.
        /// </summary>
        public string ComponentName { get { return _qualifiedName.QualifiedModuleName.ComponentName; } }

        private readonly string _parentScope;
        /// <summary>
        /// Gets the parent scope of the declaration.
        /// </summary>
        public string ParentScope { get { return _parentScope; } }

        private readonly Declaration _parentScopeDeclaration;
        /// <summary>
        /// Gets the <see cref="Declaration"/> object representing the parent scope of this declaration.
        /// </summary>
        public Declaration ParentScopeDeclaration { get { return _parentScopeDeclaration; } }

        private readonly string _identifierName;
        /// <summary>
        /// Gets the declared name of the identifier.
        /// </summary>
        public string IdentifierName { get { return _identifierName; } }

        private readonly string _asTypeName;
        /// <summary>
        /// Gets the name of the declared type.
        /// </summary>
        /// <remarks>
        /// This value is <c>null</c> if not applicable, 
        /// and <c>Variant</c> if applicable but unspecified.
        /// </remarks>
        public string AsTypeName { get { return _asTypeName; } }

        public string AsTypeNameWithoutArrayDesignator
        {
            get
            {
                if (string.IsNullOrWhiteSpace(AsTypeName))
                {
                    return AsTypeName;
                }
                return AsTypeName.Replace("(", "").Replace(")", "").Trim();
            }
        }

        public bool AsTypeIsBaseType
        {
            get
            {
                return string.IsNullOrWhiteSpace(AsTypeName) || SymbolList.BaseTypes.Contains(_asTypeName.ToUpperInvariant());
            }
        }

        private Declaration _asTypeDeclaration;
        public Declaration AsTypeDeclaration
        {
            get { return _asTypeDeclaration; }
            internal set
            {
                _asTypeDeclaration = value;
                IsSelfAssigned = _isSelfAssigned || (DeclarationType == DeclarationType.Variable &&
                                 AsTypeDeclaration.DeclarationType == DeclarationType.UserDefinedType);
            }
        }

        private readonly IReadOnlyList<DeclarationType> _neverArray = new[]
        {
            DeclarationType.ClassModule,
            DeclarationType.Control,
            DeclarationType.Document,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember,
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure,
            DeclarationType.LineLabel,
            DeclarationType.ProceduralModule,
            DeclarationType.ModuleOption,
            DeclarationType.Project,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertyLet,
            DeclarationType.UserDefinedType,
            DeclarationType.Constant
        };

        public bool IsSelected(QualifiedSelection selection)
        {
            return QualifiedName.QualifiedModuleName == selection.QualifiedName &&
                   Selection.ContainsFirstCharacter(selection.Selection);
        }

        private bool _isSelfAssigned;

        /// <summary>
        /// Gets a value indicating whether the declaration is a joined assignment (e.g. "As New xxxxx")
        /// </summary>
        public bool IsSelfAssigned
        {
            get { return _isSelfAssigned; }
            private set
            {
                _isSelfAssigned = value;
            }
        }

        private readonly Accessibility _accessibility;
        /// <summary>
        /// Gets a value specifying the declaration's visibility.
        /// This value is used in determining the declaration's scope.
        /// </summary>
        public Accessibility Accessibility { get { return _accessibility; } }

        private readonly DeclarationType _declarationType;
        /// <summary>
        /// Gets a value specifying the type of declaration.
        /// </summary>
        public DeclarationType DeclarationType { get { return _declarationType; } }

        private readonly bool _isWithEvents;
        /// <summary>
        /// Gets a value specifying whether the declared type is an event provider.
        /// </summary>
        /// <remarks>
        /// WithEvents declarations are used to identify event handler procedures in a module.
        /// </remarks>
        public bool IsWithEvents { get { return _isWithEvents; } }

        /// <summary>
        /// Returns a string representing the scope of an identifier.
        /// </summary>
        public string Scope
        {
            get
            {
                switch (_declarationType)
                {
                    case DeclarationType.Project:
                        return "VBE";
                    case DeclarationType.ClassModule:
                    case DeclarationType.ProceduralModule:
                        return _qualifiedName.QualifiedModuleName.ToString();
                    case DeclarationType.Procedure:
                    case DeclarationType.Function:
                    case DeclarationType.PropertyGet:
                    case DeclarationType.PropertyLet:
                    case DeclarationType.PropertySet:
                        return _qualifiedName.QualifiedModuleName + "." + _identifierName;
                    case DeclarationType.Event:
                        return _parentScope + "." + _identifierName;
                    default:
                        return _parentScope;
                }
            }
        }

        private readonly string _customFolder;
        private readonly bool _undeclared;
        /// <summary>
        /// Indicates whether the declaration is an ad-hoc declaration created by the resolver.
        /// </summary>
        public bool IsUndeclared { get { return _undeclared; } }

        public string CustomFolder
        {
            get
            {
                return _customFolder;
            }
        }

        public bool Equals(Declaration other)
        {
            return other != null
                && other.ProjectId == ProjectId
                && other.IdentifierName == IdentifierName
                && other.DeclarationType == DeclarationType
                && other.Scope == Scope
                && other.ParentScope == ParentScope
                && other.Selection.Equals(Selection);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as Declaration);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash * 23 + QualifiedName.QualifiedModuleName.GetHashCode();
                hash = hash * 23 + _identifierName.GetHashCode();
                hash = hash * 23 + _declarationType.GetHashCode();
                hash = hash * 23 + Scope.GetHashCode();
                hash = hash * 23 + _parentScope.GetHashCode();
                hash = hash * 23 + _selection.GetHashCode();
                return hash;
            }
        }

        public void ClearReferences()
        {
            _references = new ConcurrentBag<IdentifierReference>();
        }

        public void RemoveReferencesFrom(ICollection<QualifiedModuleName> modulesByWhichToRemoveReferences)
        {
            //This gets replaced with a new ConcurrentBag because one cannot remove specific items from a ConcurrentBag.
            //Moreover, changing to a ConcurrentDictionary<IdentifierReference,byte> breaks all sorts of tests, for some obscure reason. 
            var newReferences = new ConcurrentBag<IdentifierReference>(_references.Where(reference => !modulesByWhichToRemoveReferences.Contains(reference.QualifiedModuleName)));
        }
    }
}
