using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributeParser
    {
        IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(VBComponent component);
    }
}