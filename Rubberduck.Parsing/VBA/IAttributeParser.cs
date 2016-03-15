using System.Collections.Generic;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        IDictionary<string, IEnumerable<string>> Parse(VBComponent component);
    }
}