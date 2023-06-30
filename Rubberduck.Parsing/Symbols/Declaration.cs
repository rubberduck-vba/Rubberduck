using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Annotations.Concrete;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines a declared identifier.
    /// </summary>
    [DebuggerDisplay("({DeclarationType}) {Accessibility} {IdentifierName} As {AsTypeName} | {Selection}")]
    public class Declaration : IEquatable<Declaration>
    {
        public const int MaxModuleNameLength = 31;
        public const int MaxMemberNameLength = 255;

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
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            bool isUserDefined = true,
            IEnumerable<IParseTreeAnnotation> annotations = null,
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
                attributesPassContext,
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
            IEnumerable<IParseTreeAnnotation> annotations = null,
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
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            bool isUserDefined = true,
            IEnumerable<IParseTreeAnnotation> annotations = null,
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
            AttributesPassContext = attributesPassContext;
            IsUserDefined = isUserDefined;
            _annotations = annotations;
            Attributes = attributes ?? new Attributes();

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
                null,
                Selection.Home,
                false,
                null,
                false,
                null,
                new Attributes()) { }

        public static Declaration GetModuleParent(Declaration declaration)
        {
            if (declaration == null)
            {
                return null;
            }
            if (declaration.DeclarationType.HasFlag(DeclarationType.ClassModule) || declaration.DeclarationType == DeclarationType.ProceduralModule)
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
        public QualifiedModuleName QualifiedModuleName => QualifiedName.QualifiedModuleName;

        public ParserRuleContext Context { get; }
        public ParserRuleContext AttributesPassContext { get; }

        private ConcurrentDictionary<IdentifierReference, int> _references = new ConcurrentDictionary<IdentifierReference, int>();
        public IEnumerable<IdentifierReference> References => _references.Keys;

        protected IEnumerable<IParseTreeAnnotation> _annotations;
        public IEnumerable<IParseTreeAnnotation> Annotations => _annotations ?? new List<IParseTreeAnnotation>();

        public Attributes Attributes { get; }

        /// <summary>
        /// Gets an attribute value that contains the docstring for a member.
        /// </summary>
        public string DescriptionString
        {
            get
            {
                string literalDescription;

                var memberAttribute = Attributes.SingleOrDefault(a => 
                    a.Name == Attributes.MemberAttributeName("VB_Description", IdentifierName) || 
                    a.Name == Attributes.MemberAttributeName("VB_VarDescription", IdentifierName));

                if (memberAttribute != null)
                {
                    literalDescription = memberAttribute.Values.SingleOrDefault() ?? string.Empty;
                    return CorrectlyFormatedDescription(literalDescription);
                }

                var moduleAttribute = Attributes.SingleOrDefault(a => a.Name == "VB_Description");
                if (moduleAttribute != null)
                {
                    literalDescription = moduleAttribute.Values.SingleOrDefault() ?? string.Empty;
                    return CorrectlyFormatedDescription(literalDescription);
                }

                // fallback to description annotation; enables descriptions in document modules and non-synchronized members.
                var descriptionAnnotation = Annotations.SingleOrDefault(a =>
                    a.Annotation.GetType() == typeof(DescriptionAnnotation)
                    || a.Annotation.GetType() == typeof(VariableDescriptionAnnotation)
                    || a.Annotation.GetType() == typeof(ModuleDescriptionAnnotation));

                if (descriptionAnnotation != null)
                {
                    literalDescription = descriptionAnnotation.AnnotationArguments.FirstOrDefault();
                    return CorrectlyFormatedDescription(literalDescription);
                }
                return string.Empty;
            }
        }

        private static string CorrectlyFormatedDescription(string literalDescription)
        {
            if (string.IsNullOrEmpty(literalDescription) 
                || literalDescription.Length < 2 
                || literalDescription[0] != '"'
                || literalDescription[literalDescription.Length -1] != '"')
            {
                return literalDescription;
            }

            var text = literalDescription.Substring(1, literalDescription.Length - 2);
            return text.Replace("\"\"", "\"");
        }


        /// <summary>
        /// Gets an attribute value indicating whether a member is an enumerator provider.
        /// Types with such a member support For Each iteration.
        /// </summary>
        public bool IsEnumeratorMember => Attributes.Any(a => a.Name.EndsWith("VB_UserMemId") && a.Values.Contains("-4"));

        public virtual bool IsObject => !IsArray && IsObjectOrObjectArray;

        public virtual bool IsObjectArray => IsArray && IsObjectOrObjectArray;

        private bool IsObjectOrObjectArray
        {
            get
            {
                if (AsTypeName == Tokens.Object 
                    || (AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? false))
                {
                    return true;
                }

                var isIntrinsic = AsTypeIsBaseType
                                  || (AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false)
                                  || (AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.Enumeration) ?? false);

                return !isIntrinsic;
            }
        }

        public IdentifierReference AddReference(
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            ParserRuleContext callSiteContext,
            string identifier,
            Declaration callee,
            Selection selection,
            IEnumerable<IParseTreeAnnotation> annotations,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false,
            bool isSetAssigned = false,
            bool isIndexedDefaultMemberAccess = false,
            bool isNonIndexedDefaultMemberAccess = false,
            int defaultMemberRecursionDepth = 0,
            bool isArrayAccess = false,
            bool isProcedureCoercion = false,
            bool isInnerRecursiveDefaultMemberAccess = false,
            IdentifierReference qualifyingReference = null,
            bool isReDim = false
            )
        {
            var oldReference = _references.FirstOrDefault(r =>
                r.Key.QualifiedModuleName == module &&
                // ReSharper disable once PossibleUnintendedReferenceComparison
                r.Key.ParentScoping == scope &&
                // ReSharper disable once PossibleUnintendedReferenceComparison
                r.Key.ParentNonScoping == parent &&
                r.Key.IdentifierName == identifier &&
                r.Key.Selection == selection);
            if (oldReference.Key != null)
            {
                _references.TryRemove(oldReference.Key, out _);
            }

            var newReference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                annotations,
                isSetAssigned,
                isIndexedDefaultMemberAccess,
                isNonIndexedDefaultMemberAccess,
                defaultMemberRecursionDepth,
                isArrayAccess,
                isProcedureCoercion,
                isInnerRecursiveDefaultMemberAccess,
                qualifyingReference,
                isReDim);
            _references.AddOrUpdate(newReference, 1, (key, value) => 1);
            return newReference;
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
        /// Gets the name of the declared type as specified in code.
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
                return AsTypeName.Replace("(", string.Empty).Replace(")", string.Empty).Trim();
            }
        }

        /// <summary>
        /// Gets the fully qualified name of the declared type.
        /// </summary>
        /// <remarks>
        /// This value is <c>null</c> if not applicable, 
        /// and <c>Variant</c> if applicable but unspecified.
        /// </remarks>
        public string FullAsTypeName
        {
            get
            {
                if (AsTypeDeclaration == null)
                {
                    return AsTypeName;
                }

                if (AsTypeDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule))
                {
                    return AsTypeDeclaration.QualifiedModuleName.ToString();
                }

                //Enums and UDTs have to be qualified by the module they are contained in.
                return AsTypeDeclaration.QualifiedName.ToString();
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
                                 AsTypeDeclaration?.DeclarationType == DeclarationType.UserDefinedType);
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
                    case DeclarationType.Document:
                    case DeclarationType.ProceduralModule:
                        return QualifiedModuleName.ToString();
                    case DeclarationType.Procedure:
                    case DeclarationType.Function:
                        return $"{QualifiedModuleName}.{IdentifierName}";
                    case DeclarationType.PropertyGet:
                        return $"{QualifiedModuleName}.{IdentifierName}.Get";
                    case DeclarationType.PropertyLet:
                        return $"{QualifiedModuleName}.{IdentifierName}.Let";
                    case DeclarationType.PropertySet:
                        return $"{QualifiedModuleName}.{IdentifierName}.Set";
                    case DeclarationType.Event:
                        return $"{ParentScope}.{IdentifierName}";
                    default:
                        return ParentScope;
                }
            }
        }

        /// <summary>
        /// Indicates whether the declaration is an ad-hoc declaration created by the resolver.
        /// </summary>
        public bool IsUndeclared { get; }

        public virtual string CustomFolder => ParentDeclaration?.CustomFolder ?? ProjectName;

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

        public virtual void ClearReferences()
        {
            _references = new ConcurrentDictionary<IdentifierReference, int>();
        }

        public virtual void RemoveReferencesFrom(IReadOnlyCollection<QualifiedModuleName> modulesByWhichToRemoveReferences)
        {
            _references = new ConcurrentDictionary<IdentifierReference, int>(_references.Where(reference => !modulesByWhichToRemoveReferences.Contains(reference.Key.QualifiedModuleName)));
        }
    }
}

