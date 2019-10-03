using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProceduralModuleDeclaration : ModuleDeclaration
    {
        public ProceduralModuleDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            bool isUserDefined,
            IEnumerable<IParseTreeAnnotation> annotations,
            Attributes attributes)
            : base(
                  qualifiedName,
                  projectDeclaration,
                  name,
                  DeclarationType.ProceduralModule,
                  isUserDefined,
                  annotations,
                  attributes)
        { }

        public ProceduralModuleDeclaration(
            ComModule statics, 
            Declaration parent, 
            QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(statics.Name),
                parent,
                statics.Name,
                false,
                new List<IParseTreeAnnotation>(),
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
                new List<IParseTreeAnnotation>(),
                new Attributes()) { }

        public ProceduralModuleDeclaration(ComStruct pseudo, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName($"_{pseudo.Name}"),
                parent,
                $"_{pseudo.Name}",
                false,
                new List<IParseTreeAnnotation>(),
                new Attributes()) { }

        public bool IsPrivateModule { get; internal set; }
    }
}
