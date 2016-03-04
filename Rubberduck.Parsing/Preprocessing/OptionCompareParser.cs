using Microsoft.Vbe.Interop;
using System;
using System.Linq;

namespace Rubberduck.Parsing.Preprocessing
{
    public static class OptionCompareParser
    {
        public static VBAOptionCompare Parse(CodeModule codeModule)
        {
            int declarationLines = codeModule.CountOfDeclarationLines;
            if (declarationLines == 0)
            {
                return VBAOptionCompare.Binary;
            }
            var lines = codeModule.Lines[1, declarationLines].Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            if (lines.Any(line =>
                line.Trim().ToLower().StartsWith("option compare text")
                || line.Trim().ToLower().StartsWith("option compare database")))
            {
                return VBAOptionCompare.Text;
            }
            else
            {
                return VBAOptionCompare.Binary;
            }
        }
    }
}
