using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            IEnumerable<IAnnotation> annotations,
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
