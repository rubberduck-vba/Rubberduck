using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.DisposableWrappers;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(VBComponent component);
    }
}
