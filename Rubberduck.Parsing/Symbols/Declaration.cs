namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines a declared identifier.
    /// </summary>
    public struct Declaration
    {
        public Declaration(int projectHashCode, string parentScope,
            string projectName, string componentName, string identifierName, string asTypeName,
            Accessibility accessibility, DeclarationType declarationType, Selection selection)
        {
            _projectHashCode = projectHashCode;
            _parentScope = parentScope;
            _projectName = projectName;
            _componentName = componentName;
            _identifierName = identifierName;
            _asTypeName = asTypeName;
            _accessibility = accessibility;
            _declarationType = declarationType;
            _selection = selection;
        }

        private readonly Selection _selection;
        /// <summary>
        /// Gets a <c>Selection</c> representing the position of the declaration in the code module.
        /// </summary>
        /// <remarks>
        /// Returns <c>default(Selection)</c> for module identifiers.
        /// </remarks>
        public Selection Selection { get { return _selection; } }

        private readonly int _projectHashCode;
        /// <summary>
        /// Gets an <c>int</c> representing the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public int ProjectHashCode { get { return _projectHashCode; } }

        private readonly string _projectName;
        /// <summary>
        /// Gets the name of the VBProject the declaration is made in.
        /// </summary>
        public string ProjectName { get { return _projectName; } }

        private readonly string _componentName;
        /// <summary>
        /// Gets the name of the VBComponent the declaration is made in.
        /// </summary>
        public string ComponentName { get { return _componentName; } }

        private readonly string _parentScope;
        /// <summary>
        /// Gets the parent scope of the declaration.
        /// </summary>
        public string ParentScope { get { return _parentScope; } }

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

        /// <summary>
        /// Returns a string representing the scope of an identifier.
        /// </summary>
        public string Scope
        {
            get
            {
                switch (_declarationType)
                {
                    case DeclarationType.Class:
                    case DeclarationType.Module:
                        return _projectName;
                    case DeclarationType.Procedure:
                    case DeclarationType.Function:
                    case DeclarationType.PropertyGet:
                    case DeclarationType.PropertyLet:
                    case DeclarationType.PropertySet:
                        return _projectName + "." + _componentName;
                    default:
                        return _parentScope;
                }
            }
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Declaration))
            {
                return false;
            }

            return GetHashCode() == ((Declaration)obj).GetHashCode();
        }

        public override int GetHashCode()
        {
            return string.Concat(_projectHashCode.ToString(), _projectName, _componentName, _parentScope, _identifierName).GetHashCode();
        }
    }
}
