using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public class DocObjectDeclaration : ClassModuleDeclaration
    {
        public DocObjectDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(qualifiedName,
                projectDeclaration,
                name,
                DeclarationType.DocObject,
                true,
                annotations,
                attributes,
                true,
                true,
                false)
        { }
    }
}
