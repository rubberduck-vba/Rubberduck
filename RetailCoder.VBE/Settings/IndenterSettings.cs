using System.Xml.Serialization;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class IndenterSettings : IIndenterSettings
    {
        public bool IndentEntireProcedureBody { get; set; }
        public bool IndentFirstCommentBlock { get; set; }
        public bool IndentFirstDeclarationBlock { get; set; }
        public bool AlignCommentsWithCode { get; set; }
        public bool AlignContinuations { get; set; }
        public bool IgnoreOperatorsInContinuations { get; set; }
        public bool IndentCase { get; set; }
        public bool ForceDebugStatementsInColumn1 { get; set; }
        public bool ForceCompilerStuffInColumn1 { get; set; }
        public bool IndentCompilerDirectives { get; set; }
        public bool AlignDims { get; set; }
        public int AlignDimColumn { get; set; }
        public bool EnableUndo { get; set; }
        public EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }
        public int EndOfLineCommentColumnSpaceAlignment { get; set; }
        public int IndentSpaces { get; set; }
    }
}