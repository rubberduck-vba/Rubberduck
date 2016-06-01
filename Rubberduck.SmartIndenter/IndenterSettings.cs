using System;
using System.Xml.Serialization;
using Microsoft.Win32;

namespace Rubberduck.SmartIndenter
{
    [XmlType(AnonymousType = true)]
    public class IndenterSettings : IIndenterSettings
    {
        public virtual bool IndentEntireProcedureBody { get; set; }
        public virtual bool IndentFirstCommentBlock { get; set; }
        public virtual bool IndentFirstDeclarationBlock { get; set; }
        public virtual bool AlignCommentsWithCode { get; set; }
        public virtual bool AlignContinuations { get; set; }
        public virtual bool IgnoreOperatorsInContinuations { get; set; }
        public virtual bool IndentCase { get; set; }
        public virtual bool ForceDebugStatementsInColumn1 { get; set; }
        public virtual bool ForceCompilerDirectivesInColumn1 { get; set; }
        public virtual bool IndentCompilerDirectives { get; set; }
        public virtual bool AlignDims { get; set; }
        public virtual int AlignDimColumn { get; set; }
        public virtual bool EnableUndo { get; set; }
        public virtual EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }
        public virtual int EndOfLineCommentColumnSpaceAlignment { get; set; }
        public virtual int IndentSpaces { get; set; }

        public IndenterSettings()
        {
            var tabWidth = 4;
            var reg = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\6.0\Common", false) ??
                      Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\7.0\Common", false);
            if (reg != null)
            {
                tabWidth = Convert.ToInt32(reg.GetValue("TabWidth") ?? tabWidth);
            }

            IndentEntireProcedureBody = true;
            IndentFirstCommentBlock = true;
            IndentFirstDeclarationBlock = true;
            AlignCommentsWithCode = true;
            AlignContinuations = true;
            IgnoreOperatorsInContinuations = true;
            IndentCase = false;
            ForceDebugStatementsInColumn1 = false;
            ForceCompilerDirectivesInColumn1 = false;
            IndentCompilerDirectives = true;
            AlignDims = false;
            AlignDimColumn = 15;
            EnableUndo = true;
            EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
            EndOfLineCommentColumnSpaceAlignment = 50;
            IndentSpaces = tabWidth;
        }
    }
}
