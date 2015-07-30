﻿using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigations
{
    public interface IFindAllImplementations
    {
        void Find();
        void Find(Declaration target);
        void Find(Declaration target, VBProjectParseResult parseResult);
    }
}