using System.Xml.Serialization;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class IndenterSettings : IIndenterSettings
    {
        public bool IndentProcedure { get; set; }
        public bool IndentComment { get; set; }
        public bool IndentCase { get; set; }
        public bool IndentDim { get; set; }
        public bool AlignContinuations { get; set; }
        public bool IndentFirst { get; set; }
        public bool AlignEndOfLine { get; set; }
        public bool AlignDim { get; set; }
        public bool ForceDebugStatementsInColumn1 { get; set; }
        public bool EnableUndo { get; set; }
        public int IndentSpaces { get; set; }
        public int EndOfLineAlignColumn { get; set; }
        public int AlignDimColumn { get; set; }
        public bool ForceCompilerStuffInColumn1 { get; set; }
        public bool IndentCompilerStuff { get; set; }
        public bool AlignIgnoreOps { get; set; }
        public bool EnableIndentProcedureHotKey { get; set; }
        public bool EnableIndentModuleHotKey { get; set; }
        public EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }
    }
}