using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class ComInformation
    {
        public ComInformation(TYPEATTR typeAttributes, IMPLTYPEFLAGS implTypeFlags, ITypeInfo typeInfo, string typeName, QualifiedModuleName typeModuleName, Declaration moduleDeclaration, DeclarationType typeDeclarationType)
        {
            TypeAttributes = typeAttributes;
            ImplTypeFlags = implTypeFlags;
            TypeInfo = typeInfo;
            TypeName = typeName;
            TypeQualifiedModuleName = typeModuleName;
            ModuleDeclaration = moduleDeclaration;
            TypeDeclarationType = typeDeclarationType;
        }

        public TYPEATTR TypeAttributes { get; internal set; }
        public IMPLTYPEFLAGS ImplTypeFlags { get; internal set; }
        public ITypeInfo TypeInfo { get; internal set; }

        public string TypeName { get; internal set; }
        public QualifiedModuleName TypeQualifiedModuleName { get; internal set; }
        public Declaration ModuleDeclaration { get; internal set; }
        public DeclarationType TypeDeclarationType { get; internal set; }

        public override string ToString()
        {
            return ModuleDeclaration.IdentifierName;
        }
    }
}