using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public class ActiveXDesignerDeclaration : ClassModuleDeclaration
    {
        public ActiveXDesignerDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(qualifiedName,
                projectDeclaration,
                name,
                DeclarationType.ActiveXDesigner,
                true,
                annotations,
                attributes,
                true,
                true,
                false)
        { }
    }
}
