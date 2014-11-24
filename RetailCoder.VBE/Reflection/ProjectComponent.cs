using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Reflection
{
    [ComVisible(false)]
    public class ProjectComponent
    {
        private readonly CodeModule _module;

        public ProjectComponent(CodeModule module)
        {
            _module = module;
        }

        public IEnumerable<Member> GetMembers()
        {
            var currentLine = _module.CountOfDeclarationLines + 1;
            while (currentLine < _module.CountOfLines)
            {
                vbext_ProcKind kind;
                var name = _module.get_ProcOfLine(currentLine, out kind);

                var startLine = _module.get_ProcStartLine(name, kind);
                var lineCount = _module.get_ProcCountLines(name, kind);

                var body = _module.Lines[startLine, lineCount].Split('\n');

                Member member;
                if (Member.TryParse(body, _module.Parent.Collection.Parent.Name, _module.Parent.Name, out member))
                {
                    yield return member;
                    currentLine = startLine + lineCount;
                }
            }
        }

        private readonly IDictionary<string, MemberAttribute> _attributes;
        public bool HasAttribute(string name)
        {
            return _attributes.ContainsKey(name);
        }
    }
}
