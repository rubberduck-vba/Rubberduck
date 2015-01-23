using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;

namespace Rubberduck.Extensions
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

        /// <summary>
        /// Returns the proper file extension for the Component Type.
        /// </summary>
        /// <remarks>Document classes should properly have a ".cls" file extension.
        /// However, because they cannot be removed and imported like other component types, we need to make a distinction.</remarks>
        /// <param name="componentType"></param>
        /// <returns>File extension that includes a preceeding "dot" (.) </returns>
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
                    // documents should technically be a ".cls", but we need to be able to tell them apart.
                    return ".doccls";
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                default:
                    return string.Empty;
            }
        }
    }
}