using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ClassModuleDeclaration : Declaration
    {
        private readonly List<string> _supertypeNames;
        private readonly HashSet<Declaration> _supertypes;
        private readonly HashSet<Declaration> _subtypes;

        public ClassModuleDeclaration(
                  QualifiedMemberName qualifiedName,
                  Declaration projectDeclaration,
                  string name,
                  bool isUserDefined,
                  IEnumerable<IAnnotation> annotations,
                  Attributes attributes,
                  bool hasDefaultInstanceVariable = false,
                  bool isControl = false)
            : base(
                  qualifiedName,
                  projectDeclaration,
                  projectDeclaration,
                  name,
                  null,
                  false,
                  false,
                  Accessibility.Public,
                  DeclarationType.ClassModule,
                  null,
                  Selection.Home,
                  false,
                  null,
                  isUserDefined,
                  annotations,
                  attributes)
        {
            if (hasDefaultInstanceVariable)
            {
                _hasPredeclaredId = true;
            }
            _supertypeNames = new List<string>();
            _supertypes = new HashSet<Declaration>();
            _subtypes = new HashSet<Declaration>();
            IsControl = isControl;
        }

        // skip IDispatch.. just about everything implements it and RD doesn't need to care about it; don't care about IUnknown either
        private static readonly HashSet<string> IgnoredInterfaces = new HashSet<string>(new[] { "IDispatch", "IUnknown" });

        public ClassModuleDeclaration(ComCoClass coClass, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : base(
                module.QualifyMemberName(coClass.Name),
                parent,
                parent,
                coClass.Name,
                null,
                false,
                coClass.EventInterfaces.Any(),
                Accessibility.Public,
                DeclarationType.ClassModule,
                null,
                Selection.Home,
                false,
                null,
                false,
                new List<IAnnotation>(),
                attributes)
        {
            _supertypeNames =
                coClass.ImplementedInterfaces.Where(i => !i.IsRestricted && !IgnoredInterfaces.Contains(i.Name))
                    .Select(i => i.Name)
                    .ToList();
            _supertypes = new HashSet<Declaration>();
            _subtypes = new HashSet<Declaration>();
            IsControl = coClass.IsControl;
        }

        public ClassModuleDeclaration(ComInterface @interface, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(@interface.Name),
                parent,
                @interface.Name,
                false,
                new List<IAnnotation>(),
                attributes)
        { }

        public static IEnumerable<Declaration> GetSupertypes(Declaration type)
        {
            if (type.DeclarationType != DeclarationType.ClassModule)
            {
                return new List<Declaration>();
            }
            var classType = type as ClassModuleDeclaration;
            return classType != null ? classType.Supertypes : new List<Declaration>();
        }

        public static bool HasDefaultMember(Declaration type)
        {
            var classModule = type as ClassModuleDeclaration;
            return classModule?.DefaultMember != null;
        }

        private bool? _isExtensible;
        public bool IsExtensible 
        {
            get
            {
                AttributeNode node;
                return _isExtensible ?? (_isExtensible = Attributes.HasExtensibleAttribute(out node)).Value;
            }
        }

        private bool? _isExposed;
        /// <summary>
        /// Gets an attribute value indicating whether a class is exposed to other projects.
        /// If this value is false, any public types and members cannot be accessed from outside the project they're declared in.
        /// </summary>
        public bool IsExposed
        {
            get
            {
                AttributeNode node;
                return _isExposed ?? (_isExposed = (!IsUserDefined && IsExposedForBuiltInModules) || Attributes.HasExposedAttribute(out node)).Value;
            }
        }

        // TODO: This should only be a boolean in VBA ('Private' (false) and 'PublicNotCreatable' (true)) . For VB6 it will also need to support
        // 'SingleUse', 'GlobalSingleUse', 'MultiUse', and 'GlobalMultiUse'. See https://msdn.microsoft.com/en-us/library/aa234184%28v=vs.60%29.aspx
        // All built-ins are public (by definition).
        private static bool IsExposedForBuiltInModules { get; } = true;

        public bool IsControl { get; private set; }

        private bool? _isGlobal;
        public bool IsGlobalClassModule
        {
            get
            {
                AttributeNode node;
                return _isGlobal ?? (_isGlobal = Attributes.HasGlobalAttribute(out node) || IsGlobalFromSubtypes()).Value;
            }
        }

        private bool IsGlobalFromSubtypes()
        {
            return Subtypes.Any(subtype => subtype is ClassModuleDeclaration && ((ClassModuleDeclaration)subtype).IsGlobalClassModule);
        }

        private bool? _hasPredeclaredId;
        /// <summary>
        /// Gets an attribute value indicating whether a class has a predeclared ID.
        /// Such classes can be treated as "static classes", or as far as resolving is concerned, as standard modules.
        /// </summary>
        public bool HasPredeclaredId
        {
            get
            {
                AttributeNode node;
                return _hasPredeclaredId ?? (_hasPredeclaredId = Attributes.HasPredeclaredIdAttribute(out node)).Value;
            }
        }

        public bool HasDefaultInstanceVariable => HasPredeclaredId || IsGlobalClassModule;

        public Declaration DefaultMember { get; internal set; }

        public IReadOnlyList<string> SupertypeNames => _supertypeNames;

        public IReadOnlyList<Declaration> Supertypes => _supertypes.ToList();

        public IReadOnlyList<Declaration> Subtypes => _subtypes.ToList();

        public void AddSupertype(string supertypeName)
        {
            _supertypeNames.Add(supertypeName);
        }

        public void AddSupertype(Declaration supertype)
        {
            _supertypes.Add(supertype);
        }

        public void AddSubtype(Declaration subtype)
        {
            _subtypes.Add(subtype);
        }
    }
}
