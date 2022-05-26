using System.Collections.Generic;

namespace Rubberduck.SmartIndenter
{
    public interface ISimpleIndenter
    {
        IEnumerable<string> Indent(string code, IIndenterSettings settings = null);
        IEnumerable<string> Indent(IEnumerable<string> lines, IIndenterSettings settings = null);
        IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines, IIndenterSettings settings = null);
    }
}
