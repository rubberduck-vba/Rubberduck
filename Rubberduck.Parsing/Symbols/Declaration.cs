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
            bool isUserDefined = true,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null,
            bool undeclared = false)
            : this(
                qualifiedName,
                parentDeclaration,
                parentScope?.Scope,
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
                isUserDefined,
                annotations,
                attributes)
        {
            ParentScopeDeclaration = parentScope;
            IsUndeclared = undeclared;
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
            bool isUserDefined = true,
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
                  isUserDefined,
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
            bool isUserDefined = true,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null)
        {
            QualifiedName = qualifiedName;            
            ParentDeclaration = parentDeclaration;
            ParentScopeDeclaration = ParentDeclaration;
            ParentScope = parentScope ?? string.Empty;
            IdentifierName = qualifiedName.MemberName;
            AsTypeName = asTypeName;
            IsSelfAssigned = isSelfAssigned;
            IsWithEvents = isWithEvents;
            Accessibility = accessibility;
            DeclarationType = declarationType;
            Selection = selection;
            Context = context;
            IsUserDefined = isUserDefined;
            _annotations = annotations;
            _attributes = attributes ?? new Attributes();

            ProjectId = QualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = GetProjectParent(parentDeclaration);
            if (projectDeclaration != null)
            {
                ProjectName = projectDeclaration.IdentifierName;
            }
            else if (DeclarationType == DeclarationType.Project)
            {
                ProjectName = IdentifierName;
            }

            CustomFolder = FolderFromAnnotations();
            IsArray = isArray;
            AsTypeContext = asTypeContext;
            TypeHint = typeHint;
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
            false,
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
                false,
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
                null,
                false,
                null,
                new Attributes()) { }

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
                null,
                false,
                null,
                new Attributes()) { }

        private string FolderFromAnnotations()
            {
                var @namespace = Annotations.FirstOrDefault(annotation => annotation.AnnotationType == AnnotationType.Folder);
                string result;
                if (@namespace == null)
                {
                    result = string.IsNullOrEmpty(QualifiedName.QualifiedModuleName.ProjectName)
                        ? ProjectId
                        : QualifiedName.QualifiedModuleName.ProjectName;
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

        public bool IsArray { get; }
        public VBAParser.AsTypeClauseContext AsTypeContext { get; }
        public string TypeHint { get; }
        public bool HasTypeHint => !string.IsNullOrWhiteSpace(TypeHint);

        public bool IsTypeSpecified => HasTypeHint || AsTypeContext != null;

        public bool IsUserDefined { get; }

        public Declaration ParentDeclaration { get; }

        public QualifiedMemberName QualifiedName { get; }

        public ParserRuleContext Context { get; }

        private ConcurrentBag<IdentifierReference> _references = new ConcurrentBag<IdentifierReference>();
        public IEnumerable<IdentifierReference> References => _references;

        private readonly IEnumerable<IAnnotation> _annotations;
        public IEnumerable<IAnnotation> Annotations => _annotations ?? new List<IAnnotation>();

        private readonly Attributes _attributes;
        public IReadOnlyDictionary<string, IEnumerable<string>> Attributes => _attributes;

        /// <summary>
        /// Gets an attribute value that contains the docstring for a member.
        /// </summary>
        public string DescriptionString
        {
            get
            {
                IEnumerable<string> value;
                if (_attributes.TryGetValue($"{IdentifierName}.VB_Description", out value))
                {
                    return value.Single();
                }
                if(_attributes.TryGetValue("VB_Description", out value))
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

        /// <summary>
        /// Gets a <c>Selection</c> representing the position of the declaration in the code module.
        /// </summary>
        /// <remarks>
        /// Returns <c>default(Selection)</c> for module identifiers.
        /// </remarks>
        public Selection Selection { get; }

        public QualifiedSelection QualifiedSelection => new QualifiedSelection(QualifiedName.QualifiedModuleName, Selection);

        /// <summary>
        /// Gets a reference to the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public virtual IVBProject Project => ParentDeclaration.Project;

        /// <summary>
        /// Gets a unique identifier for the VBProject the declaration is made in.
        /// </summary>
        public string ProjectId { get; }

        public string ProjectName { get; }

        /// <summary>
        /// WARNING: This property has side effects. It changes the ActiveVBProject, which causes a flicker in the VBE.
        /// This should only be called if it is *absolutely* necessary.
        /// </summary>
        public virtual string ProjectDisplayName => ParentDeclaration.ProjectDisplayName;

        public object[] ToArray()
        {
            return new object[] { ProjectName, CustomFolder, ComponentName, DeclarationType.ToString(), Scope, IdentifierName, AsTypeName };
        }


        /// <summary>
        /// Gets the name of the VBComponent the declaration is made in.
        /// </summary>
        public string ComponentName => QualifiedName.QualifiedModuleName.ComponentName;

        /// <summary>
        /// Gets the parent scope of the declaration.
        /// </summary>
        public string ParentScope { get; }

        /// <summary>
        /// Gets the <see cref="Declaration"/> object representing the parent scope of this declaration.
        /// </summary>
        public Declaration ParentScopeDeclaration { get; }

        /// <summary>
        /// Gets the declared name of the identifier.
        /// </summary>
        public string IdentifierName { get; }

        /// <summary>
        /// Gets the name of the declared type.
        /// </summary>
        /// <remarks>
        /// This value is <c>null</c> if not applicable, 
        /// and <c>Variant</c> if applicable but unspecified.
        /// </remarks>
        public string AsTypeName { get; }

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

        public bool AsTypeIsBaseType => string.IsNullOrWhiteSpace(AsTypeName) || SymbolList.BaseTypes.Contains(AsTypeName.ToUpperInvariant());

        private Declaration _asTypeDeclaration;
        public Declaration AsTypeDeclaration
        {
            get { return _asTypeDeclaration; }
            internal set
            {
                _asTypeDeclaration = value;
                IsSelfAssigned = IsSelfAssigned || (DeclarationType == DeclarationType.Variable &&
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

        /// <summary>
        /// Gets a value indicating whether the declaration is a joined assignment (e.g. "As New xxxxx")
        /// </summary>
        public bool IsSelfAssigned { get; private set; }

        /// <summary>
        /// Gets a value specifying the declaration's visibility.
        /// This value is used in determining the declaration's scope.
        /// </summary>
        public Accessibility Accessibility { get; }

        /// <summary>
        /// Gets a value specifying the type of declaration.
        /// </summary>
        public DeclarationType DeclarationType { get; }

        /// <summary>
        /// Gets a value specifying whether the declared type is an event provider.
        /// </summary>
        /// <remarks>
        /// WithEvents declarations are used to identify event handler procedures in a module.
        /// </remarks>
        public bool IsWithEvents { get; }

        /// <summary>
        /// Returns a string representing the scope of an identifier.
        /// </summary>
        public string Scope
        {
            get
            {
                switch (DeclarationType)
                {
                    case DeclarationType.Project:
                        return "VBE";
                    case DeclarationType.ClassModule:
                    case DeclarationType.ProceduralModule:
                        return QualifiedName.QualifiedModuleName.ToString();
                    case DeclarationType.Procedure:
                    case DeclarationType.Function:
                    case DeclarationType.PropertyGet:
                    case DeclarationType.PropertyLet:
                    case DeclarationType.PropertySet:
                        return QualifiedName.QualifiedModuleName + "." + IdentifierName;
                    case DeclarationType.Event:
                        return ParentScope + "." + IdentifierName;
                    default:
                        return ParentScope;
                }
            }
        }

        /// <summary>
        /// Indicates whether the declaration is an ad-hoc declaration created by the resolver.
        /// </summary>
        public bool IsUndeclared { get; }

        public string CustomFolder { get; }

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
                hash = hash * 23 + IdentifierName.GetHashCode();
                hash = hash * 23 + DeclarationType.GetHashCode();
                hash = hash * 23 + Scope.GetHashCode();
                hash = hash * 23 + ParentScope.GetHashCode();
                hash = hash * 23 + Selection.GetHashCode();
                return hash;
            }
        }

        public void ClearReferences()
        {
            _references = new ConcurrentBag<IdentifierReference>();
        }

        public void RemoveReferencesFrom(IReadOnlyCollection<QualifiedModuleName> modulesByWhichToRemoveReferences)
        {
            //This gets replaced with a new ConcurrentBag because one cannot remove specific items from a ConcurrentBag.
            //Moreover, changing to a ConcurrentDictionary<IdentifierReference,byte> breaks all sorts of tests, for some obscure reason. 
            _references = new ConcurrentBag<IdentifierReference>(_references.Where(reference => !modulesByWhichToRemoveReferences.Contains(reference.QualifiedModuleName)));
        }
    }
}
