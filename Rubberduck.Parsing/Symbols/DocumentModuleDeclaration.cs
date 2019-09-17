using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class DocumentModuleDeclaration : ClassModuleDeclaration
    {
        public DocumentModuleDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            IEnumerable<IParseTreeAnnotation> annotations,
            Attributes attributes)
            : base(qualifiedName, 
                projectDeclaration,
                name,
                true,
                annotations,
                attributes,
                true,
                true,
                false,
                true)
        { }
    }
}
