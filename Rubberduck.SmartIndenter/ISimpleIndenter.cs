using System.Collections.Generic;

namespace Rubberduck.SmartIndenter
{
    public interface ISimpleIndenter
    {
        IEnumerable<string> Indent(string code);
        IEnumerable<string> Indent(IEnumerable<string> lines);
        IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines);
    }
}
