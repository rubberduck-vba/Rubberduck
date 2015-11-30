using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenterSettings
    {
        bool IndentProcedure { get; set; }
        bool IndentComment { get; set; }
        bool IndentCase { get; set; }
        bool IndentDim { get; set; }
        bool AlignContinuations { get; set; }
        bool IndentFirst { get; set; }
        bool AlignEndOfLine { get; set; }
        bool AlignDim { get; set; }
        bool DebugColumn1 { get; set; }
        bool EnableUndo { get; set; }
        int IndentSpaces { get; set; }
        int EndOfLineAlignColumn { get; set; }
        int AlignDimColumn { get; set; }
        bool CompilerStuffColumn1 { get; set; }
        bool IndentCompilerStuff { get; set; }
        bool AlignIgnoreOps { get; set; }
    }
}
