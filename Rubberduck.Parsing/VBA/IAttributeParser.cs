using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(VBComponent component);
    }
}
