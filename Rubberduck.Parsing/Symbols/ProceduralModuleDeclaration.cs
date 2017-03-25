using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
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
            bool isUserDefined,
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
                  isUserDefined,
                  annotations,
                  attributes) { }

        public ProceduralModuleDeclaration(ComModule statics, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(statics.Name),
                parent,
                statics.Name,
                false,
                new List<IAnnotation>(),
                attributes)
        {
            IsPrivateModule = statics.IsRestricted;
        }

        //These are the pseudo-module ctor for COM enumerations and types.
        public ProceduralModuleDeclaration(ComEnumeration pseudo, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName($"_{pseudo.Name}"),
                parent,
                $"_{pseudo.Name}",
                false,
                new List<IAnnotation>(),
                new Attributes()) { }

        public ProceduralModuleDeclaration(ComStruct pseudo, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName($"_{pseudo.Name}"),
                parent,
                $"_{pseudo.Name}",
                false,
                new List<IAnnotation>(),
                new Attributes()) { }

        public bool IsPrivateModule { get; internal set; }
    }
}
