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
                  bool isBuiltIn,
                  IEnumerable<IAnnotation> annotations,
                  Attributes attributes,
                  bool hasDefaultInstanceVariable = false)
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
                  isBuiltIn,
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
                true,
                new List<IAnnotation>(),
                attributes)
        {
            _supertypeNames =
                coClass.ImplementedInterfaces.Where(i => !i.IsRestricted && !IgnoredInterfaces.Contains(i.Name))
                    .Select(i => i.Name)
                    .ToList();
            _supertypes = new HashSet<Declaration>();
            _subtypes = new HashSet<Declaration>();
            IsExtensible = coClass.IsExtensible;
        }

        public ClassModuleDeclaration(ComInterface intrface, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(intrface.Name),
                parent,
                intrface.Name,
                true,
                new List<IAnnotation>(),
                attributes)
        {
            IsExtensible = intrface.IsExtensible;
        }

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
            return classModule != null && classModule.DefaultMember != null;
        }

        public bool IsExtensible { get; set; }

        private bool? _isExposed;
        /// <summary>
        /// Gets an attribute value indicating whether a class is exposed to other projects.
        /// If this value is false, any public types and members cannot be accessed from outside the project they're declared in.
        /// </summary>
        public bool IsExposed
        {
            get
            {
                if (_isExposed.HasValue)
                {
                    return _isExposed.Value;
                }
                if (IsBuiltIn)
                {
                    _isExposed = IsExposedForBuiltInModules();
                    return _isExposed.Value;
                }
                _isExposed = HasAttribute("VB_Exposed");
                return _isExposed.Value;
            }
        }

            // TODO: Find out if there's info about "being exposed" in type libraries.
            // We take the conservative approach of treating all type library modules as exposed.
            private static bool IsExposedForBuiltInModules()
            {
                return true;
            }

            private bool HasAttribute(string attributeName)
            {
                var hasAttribute = false;
                IEnumerable<string> value;
                if (Attributes.TryGetValue(attributeName, out value))
                {
                    hasAttribute = value.Single() == "True";
                }
                return hasAttribute;
            }
              

        private bool? _isGlobal;
        public bool IsGlobalClassModule
        {
            get
            {
                if (_isGlobal.HasValue)
                {
                    return _isGlobal.Value;
                }
                _isGlobal = HasAttribute("VB_GlobalNamespace") || IsGlobalFromSubtypes();
                return _isGlobal.Value;
            }
        }

            private bool IsGlobalFromSubtypes()
            {
                return Subtypes.Any(subtype => (subtype is ClassModuleDeclaration && ((ClassModuleDeclaration)subtype).IsGlobalClassModule));
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
                if (_hasPredeclaredId.HasValue)
                {
                    return _hasPredeclaredId.Value;
                }
                _hasPredeclaredId = HasAttribute("VB_PredeclaredId");
                return _hasPredeclaredId.Value;
            }
        }


        public bool HasDefaultInstanceVariable
        {
            get
            {
                return HasPredeclaredId || IsGlobalClassModule;
            }
        }

        public Declaration DefaultMember { get; internal set; }

        public IReadOnlyList<string> SupertypeNames
        {
            get
            {
                return _supertypeNames;
            }
        }

        public IReadOnlyList<Declaration> Supertypes
        {
            get
            {
                return _supertypes.ToList();
            }
        }

        public IReadOnlyList<Declaration> Subtypes
        {
            get
            {
                return _subtypes.ToList();
            }
        }

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
