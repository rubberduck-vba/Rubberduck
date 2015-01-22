using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Reflection
{
    [ComVisible(false)]
    public static class VBComponentExtensions
    {
        public static bool HasAttribute<TAttribute>(this CodeModule code) where TAttribute : MemberAttributeBase, new()
        {
            return HasAttribute(code, new TAttribute().Name);
        }

        public static bool HasAttribute(this CodeModule code, string name)
        {
            if (code.CountOfDeclarationLines == 0)
            {
                return false;
            }
            var moduleAttributes = MemberAttribute.GetAttributes(code.Lines[1, code.CountOfDeclarationLines].Split('\n'));
            return (moduleAttributes.Any(attribute => attribute.Name == name));
        }

        public static IEnumerable<Member> GetMembers(this VBComponent component, vbext_ProcKind? procedureKind = null)
        {
            return GetMembers(component.CodeModule, procedureKind);
        }

        public static IEnumerable<Member> GetMembers(this CodeModule module, vbext_ProcKind? procedureKind = null)
        {
            var currentLine = module.CountOfDeclarationLines + 1;
            while (currentLine < module.CountOfLines)
            {
                vbext_ProcKind kind;
                var name = module.get_ProcOfLine(currentLine, out kind);

                if ((procedureKind ?? kind) == kind)
                {
                    var startLine = module.get_ProcStartLine(name, kind);
                    var lineCount = module.get_ProcCountLines(name, kind);

                    var body = module.Lines[startLine, lineCount].Split('\n');
                    
                    Member member;
                    if (Member.TryParse(body, module.Parent.Collection.Parent.Name, module.Parent.Name, out member))
                    {
                        yield return member;
                        currentLine = startLine + lineCount;
                    }
                    else
                    {
                        currentLine = currentLine + 1;
                    }
                }
            }
        }

        public static string FileExtension(this vbext_ComponentType componentType)
        {
            switch (componentType)
            {
                case vbext_ComponentType.vbext_ct_ClassModule:
                    return ".cls";
                case vbext_ComponentType.vbext_ct_MSForm:
                    return ".frm";
                case vbext_ComponentType.vbext_ct_StdModule:
                    return ".bas";
                case vbext_ComponentType.vbext_ct_Document:
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                default:
                    return string.Empty;
            }
        }
    }
}