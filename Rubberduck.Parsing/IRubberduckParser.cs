using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser
    {
        RubberduckParserState State { get; }

        void Parse(VBE vbe);
        void Parse(VBProject vbProject);
        Task ParseAsync(VBComponent vbComponent);
    }
}