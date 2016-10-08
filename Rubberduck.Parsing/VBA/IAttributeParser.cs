using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(IVBComponent component);
    }
}
