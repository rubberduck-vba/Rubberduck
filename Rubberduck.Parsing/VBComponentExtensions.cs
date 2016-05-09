using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Reflection;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    [ComVisible(false)]
    public static class VBComponentExtensions
    {
        public static bool HasAttribute<TAttribute>(this CodeModule code) where TAttribute : MemberAttributeBase, new()
        {
            return HasAttribute(code, new TAttribute().Name);
        }

        private static bool HasAttribute(this CodeModule code, string name)
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

        private static IEnumerable<Member> GetMembers(this CodeModule module, vbext_ProcKind? procedureKind = null)
        {
            var currentLine = module.CountOfDeclarationLines + 1;
            while (currentLine <= module.CountOfLines)
            {
                vbext_ProcKind kind;
                var name = module.get_ProcOfLine(currentLine, out kind);

                if ((procedureKind ?? kind) == kind)
                {
                    var startLine = module.get_ProcStartLine(name, kind);
                    var lineCount = module.get_ProcCountLines(name, kind);

                    var body = module.Lines[startLine, lineCount].Split('\n');

                    Member member;
                    if (Member.TryParse(body, new QualifiedModuleName(module.Parent), out member))
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
    }
}