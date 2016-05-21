using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;

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
                  null,
                  false,
                  false,
                  Accessibility.Public,
                  DeclarationType.ProceduralModule,
                  null,
                  Selection.Home,
                  false,
                  null,
                  isBuiltIn,
                  annotations,
                  attributes)
        {
        }

        public bool IsPrivateModule { get; internal set; }
    }
}
