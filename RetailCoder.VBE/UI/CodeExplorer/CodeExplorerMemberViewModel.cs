using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerMemberViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;

        private static readonly DeclarationType[] SubMemberTypes =
        {
            DeclarationType.EnumerationMember, 
            DeclarationType.UserDefinedTypeMember            
        };

        public CodeExplorerMemberViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            if (declarations != null)
            {
                _members = declarations.Where(item => SubMemberTypes.Contains(item.DeclarationType) && item.ParentDeclaration.Equals(declaration))
                                       .Select(item => new CodeExplorerMemberViewModel(item, null));
            }
        }

        public string Name { get { return _declaration.IdentifierName; } }
        //public string Signature { get { return _declaration.IdentifierName; } }

        private readonly IEnumerable<CodeExplorerMemberViewModel> _members;
        public IEnumerable<CodeExplorerMemberViewModel> Members { get { return _members; } }
    }
}