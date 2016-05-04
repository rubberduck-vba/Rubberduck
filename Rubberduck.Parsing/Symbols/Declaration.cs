using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines a declared identifier.
    /// </summary>
    [DebuggerDisplay("({DeclarationType}) {Accessibility} {IdentifierName} As {AsTypeName} | {Selection}")]
    public class Declaration : IEquatable<Declaration>
    {
        public static readonly string[] BASE_TYPES = new string[]
        {
                "BOOLEAN",
                "BYTE",
                "CURRENCY",
                "DATE",
                "DOUBLE",
                "INTEGER",
                "LONG",
                "LONGLONG",
                "LONGPTR",
                "SINGLE",
                "STRING",
                "VARIANT",
                "OBJECT"
        };

        public Declaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            Declaration parentScope,
            string asTypeName,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            DeclarationType declarationType,
            ParserRuleContext context,
            Selection selection,
            bool isBuiltIn = true,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null)
            : this(
                qualifiedName, parentDeclaration, parentScope == null ? null : parentScope.Scope, asTypeName,
                isSelfAssigned, isWithEvents,
                accessibility, declarationType, context, selection, isBuiltIn, annotations, attributes)
        {
            _parentScopeDeclaration = parentScope;
        }

        public Declaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            string parentScope,
            string asTypeName,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            DeclarationType declarationType,
            bool isBuiltIn = true,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null)
            : this(qualifiedName, parentDeclaration, parentScope, asTypeName, isSelfAssigned, isWithEvents, accessibility, declarationType, null, Selection.Home, isBuiltIn, annotations, attributes)
        { }

        public Declaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            string parentScope,
            string asTypeName,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            DeclarationType declarationType,
            ParserRuleContext context,
            Selection selection,
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
            _isSelfAssigned = isSelfAssigned || (declarationType == DeclarationType.Variable && parentDeclaration != null && parentDeclaration.IdentifierName == ComponentName);
            _isWithEvents = isWithEvents;
            _accessibility = accessibility;
            _declarationType = declarationType;
            _selection = selection;
            _context = context;
            _isBuiltIn = isBuiltIn;
            _annotations = annotations;
            _attributes = attributes ?? new Attributes();

            _projectId = _qualifiedName.QualifiedModuleName.ProjectId;

            var @namespace = Annotations.FirstOrDefault(annotation => annotation.AnnotationType == AnnotationType.Folder);
            string result;
            if (@namespace == null)
            {
                result = _qualifiedName.QualifiedModuleName.Project == null
                    ? _projectId
                    : _qualifiedName.QualifiedModuleName.Project.Name;
            }
            else
            {
                var value = ((FolderAnnotation)@namespace).FolderName;
                result = value;
            }
            _customFolder = result;
        }

        public static bool HasParameter(DeclarationType declarationType)
        {
            return
                declarationType.HasFlag(DeclarationType.Property)
                || declarationType == DeclarationType.Function
                || declarationType == DeclarationType.Procedure;
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

        private readonly bool _isBuiltIn;
        /// <summary>
        /// Marks a declaration as non-user code, e.g. the <see cref="VbaStandardLib"/> or <see cref="ExcelObjectModel"/>.
        /// </summary>
        public bool IsBuiltIn { get { return _isBuiltIn; } }

        private readonly Declaration _parentDeclaration;
        public Declaration ParentDeclaration { get { return _parentDeclaration; } }

        private readonly QualifiedMemberName _qualifiedName;
        public QualifiedMemberName QualifiedName { get { return _qualifiedName; } }

        private ParserRuleContext _context;
        public ParserRuleContext Context
        {
            get
            {
                return _context;
            }
            set
            {
                _context = value;
            }
        }

        private ConcurrentBag<IdentifierReference> _references = new ConcurrentBag<IdentifierReference>();
        public IEnumerable<IdentifierReference> References
        {
            get
            {
                return _references.Union(_memberCalls);
            }
            set
            {
                _references = new ConcurrentBag<IdentifierReference>(value);
            }
        }

        private ConcurrentBag<IdentifierReference> _memberCalls = new ConcurrentBag<IdentifierReference>();
        public IEnumerable<IdentifierReference> MemberCalls
        {
            get
            {
                return _memberCalls.ToList();
            }
            set
            {
                _memberCalls = new ConcurrentBag<IdentifierReference>(value);
            }
        }

        private readonly IEnumerable<IAnnotation> _annotations;
        public IEnumerable<IAnnotation> Annotations { get { return _annotations ?? new List<IAnnotation>(); } }

        private readonly Attributes _attributes;
        public IReadOnlyDictionary<string, IEnumerable<string>> Attributes { get { return _attributes; } }

        /// <summary>
        /// Gets an attribute value indicating whether a class has a predeclared ID.
        /// Such classes can be treated as "static classes", or as far as resolving is concerned, as standard modules.
        /// </summary>
        public bool HasPredeclaredId
        {
            get
            {
                IEnumerable<string> value;
                if (_attributes.TryGetValue("VB_PredeclaredId", out value))
                {
                    return value.Single() == "True";
                }

                return false;
            }
        }

        /// <summary>
        /// Gets an attribute value that contains the docstring for a member.
        /// </summary>
        public string DescriptionString
        {
            get
            {
                IEnumerable<string> value;
                if (_attributes.TryGetValue(IdentifierName + ".VB_Description", out value))
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
                if (_attributes.TryGetValue(IdentifierName + ".VB_UserMemId", out value))
                {
                    return value.Single() == "-4";
                }

                return false;
            }
        }

        public bool IsInspectionDisabled(string inspectionName)
        {
            return Annotations.Any(annotation =>
                annotation.AnnotationType == AnnotationType.Ignore
                && ((IgnoreAnnotation)annotation).IsIgnored(inspectionName));
        }

        public void AddReference(IdentifierReference reference)
        {
            if (reference == null || reference.Declaration.Context == reference.Context)
            {
                return;
            }
            if (reference.Context.Parent != _context
                && !_references.Select(r => r.Context).Contains(reference.Context.Parent)
                && !_references.Any(r => r.QualifiedModuleName == reference.QualifiedModuleName
                    && r.Selection.StartLine == reference.Selection.StartLine
                    && r.Selection.EndLine == reference.Selection.EndLine
                    && r.Selection.StartColumn == reference.Selection.StartColumn
                    && r.Selection.EndColumn == reference.Selection.EndColumn))
            {
                _references.Add(reference);
            }
        }

        public void AddMemberCall(IdentifierReference reference)
        {
            if (reference == null || reference.Declaration == null || reference.Declaration.Context == reference.Context)
            {
                return;
            }

            _memberCalls.Add(reference);
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
        public VBProject Project { get { return _qualifiedName.QualifiedModuleName.Project; } }

        private readonly string _projectId;
        /// <summary>
        /// Gets a unique identifier for the VBProject the declaration is made in.
        /// </summary>
        public string ProjectId { get { return _projectId; } }

        public string ProjectName
        {
            get
            {
                if (Project != null)
                {
                    return Project.Name;
                }
                // Referenced projects have their identifier name set as their project name.
                return IdentifierName;
            }
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
        
        public bool AsTypeIsBaseType
        {
            get
            {
                return string.IsNullOrWhiteSpace(AsTypeName) || BASE_TYPES.Contains(_asTypeName.ToUpperInvariant());
            }
        }

        public Declaration AsTypeDeclaration { get; internal set; }

        private bool? _isArray;

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
            DeclarationType.Constant,
        };

        public virtual bool IsArray()
        {
            if (Context == null || _neverArray.Any(item => DeclarationType.HasFlag(item)))
            {
                return false;
            }

            if (_isArray.HasValue)
            {
                return _isArray.Value;
            }

            var rParenMethod = Context.GetType().GetMethod("RPAREN");
            if (rParenMethod == null)
            {
                _isArray = false;
                return _isArray.Value;
            }

            var declaration = (dynamic)Context;
            _isArray = declaration.LPAREN() != null && declaration.RPAREN() != null;
            return _isArray.Value;
        }

        private bool? _isTypeSpecified;

        private readonly IReadOnlyList<DeclarationType> _neverSpecified = new[]
        {
            DeclarationType.Procedure,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
            DeclarationType.UserDefinedType,
            DeclarationType.ClassModule,
            DeclarationType.Control,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember,
            DeclarationType.LibraryProcedure,
            DeclarationType.LineLabel,
            DeclarationType.ModuleOption,
            DeclarationType.Project,
        };

        public VBAParser.AsTypeClauseContext GetAsTypeContext()
        {
            return (VBAParser.AsTypeClauseContext)((dynamic)Context).asTypeClause();
        }

        public virtual bool IsTypeSpecified()
        {
            if (Context == null || _neverSpecified.Any(item => DeclarationType.HasFlag(item)))
            {
                return false;
            }

            if (_isTypeSpecified.HasValue)
            {
                return _isTypeSpecified.Value;
            }

            var method = Context.GetType().GetMethod("asTypeClause");
            if (method == null)
            {
                _isTypeSpecified = false;
                return false;
            }

            if (HasTypeHint())
            {
                _isTypeSpecified = false;
                return true;
            }

            _isTypeSpecified = ((dynamic)Context).asTypeClause() is VBAParser.AsTypeClauseContext;
            return _isTypeSpecified.Value;
        }

        private bool? _hasTypeHint;

        public bool HasTypeHint()
        {
            if (_hasTypeHint.HasValue)
            {
                return _hasTypeHint.Value;
            }

            string token;
            return HasTypeHint(out token);
        }

        private readonly IReadOnlyList<DeclarationType> _neverHinted = new[]
        {
            DeclarationType.ClassModule,
            DeclarationType.LineLabel,
            DeclarationType.ModuleOption,
            DeclarationType.Project,
            DeclarationType.Control,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember,
            DeclarationType.LibraryProcedure,
            DeclarationType.Procedure,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
        };

        public bool HasTypeHint(out string token)
        {
            if (Context == null || _neverHinted.Any(item => DeclarationType.HasFlag(item)))
            {
                token = null;
                _hasTypeHint = false;
                return _hasTypeHint.Value;
            }

            try
            {
                var hint = ((dynamic)Context).typeHint();
                token = hint == null ? null : hint.GetText();
                _hasTypeHint = hint != null;
                return _hasTypeHint.Value;
            }
            catch (RuntimeBinderException)
            {
                token = null;
                _hasTypeHint = false;
                return _hasTypeHint.Value;
            }
        }

        public bool IsSelected(QualifiedSelection selection)
        {
            return QualifiedName.QualifiedModuleName == selection.QualifiedName &&
                   Selection.ContainsFirstCharacter(selection.Selection);
        }

        private readonly bool _isSelfAssigned;
        /// <summary>
        /// Gets a value indicating whether the declaration is a joined assignment (e.g. "As New xxxxx")
        /// </summary>
        public bool IsSelfAssigned { get { return _isSelfAssigned; } }

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
            while (!_references.IsEmpty)
            {
                IdentifierReference reference;
                _references.TryTake(out reference);
            }
        }
    }
}
