using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public class UserControlDeclaration : ClassModuleDeclaration
    {
        public UserControlDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(qualifiedName,
                projectDeclaration,
                name,
                DeclarationType.UserControl,
                true,
                annotations,
                attributes,
                true,
                true,
                false)
        { }
    }
}
