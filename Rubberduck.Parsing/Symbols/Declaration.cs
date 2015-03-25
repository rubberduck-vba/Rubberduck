using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines a declared identifier.
    /// </summary>
    public class Declaration
    {
        public Declaration(QualifiedMemberName qualifiedName, string parentScope,
            string identifierName, string asTypeName, bool isSelfAssigned,
            Accessibility accessibility, DeclarationType declarationType, ParserRuleContext context)
        {
            _qualifiedName = qualifiedName;
            _parentScope = parentScope;
            _identifierName = identifierName;
            _asTypeName = asTypeName;
            _isSelfAssigned = isSelfAssigned;
            _accessibility = accessibility;
            _declarationType = declarationType;
            _selection = context.GetSelection();
            _context = context;
        }

        private readonly QualifiedMemberName _qualifiedName;
        public QualifiedMemberName QualifiedName { get { return _qualifiedName; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly IList<IdentifierReference> _references = new List<IdentifierReference>();
        public IEnumerable<IdentifierReference> References { get { return _references; } }

        public void AddReference(IdentifierReference reference)
        {
            _references.Add(reference);
        }

        private readonly Selection _selection;
        /// <summary>
        /// Gets a <c>Selection</c> representing the position of the declaration in the code module.
        /// </summary>
        /// <remarks>
        /// Returns <c>default(Selection)</c> for module identifiers.
        /// </remarks>
        public Selection Selection { get { return _selection; } }

        /// <summary>
        /// Gets an <c>int</c> representing the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public int ProjectHashCode { get { return _qualifiedName.QualifiedModuleName.ProjectHashCode; } }

        /// <summary>
        /// Gets the name of the VBProject the declaration is made in.
        /// </summary>
        public string ProjectName { get { return _qualifiedName.QualifiedModuleName.ProjectName; } }

        /// <summary>
        /// Gets the name of the VBComponent the declaration is made in.
        /// </summary>
        public string ComponentName { get { return _qualifiedName.QualifiedModuleName.ModuleName; } }

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
                        return _qualifiedName.QualifiedModuleName.ProjectName + "." + _qualifiedName.QualifiedModuleName.ModuleName;
                    case DeclarationType.Procedure:
                    case DeclarationType.Function:
                    case DeclarationType.PropertyGet:
                    case DeclarationType.PropertyLet:
                    case DeclarationType.PropertySet:
                        return _qualifiedName.QualifiedModuleName.ProjectName + "." + _qualifiedName.QualifiedModuleName.ModuleName + "." + _identifierName;
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
            return string.Concat(ProjectHashCode.ToString(), ProjectName, ComponentName, _parentScope, _identifierName).GetHashCode();
        }
    }
}
