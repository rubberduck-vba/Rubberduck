using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProceduralModuleDeclaration : Declaration
    {
        public ProceduralModuleDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
                  bool isBuiltIn,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                  qualifiedName,
                  projectDeclaration,
                  projectDeclaration,
                  name,
                  false,
                  false,
                  Accessibility.Public,
                  DeclarationType.ProceduralModule,
                  null,
                  Selection.Home,
                  isBuiltIn,
                  annotations,
                  attributes)
        {
        }

        public bool IsPrivateModule { get; internal set; }
    }
}
