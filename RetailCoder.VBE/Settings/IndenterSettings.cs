//using System;
//using System.Xml.Serialization;
//using Microsoft.Win32;
//using Rubberduck.SmartIndenter;

//namespace Rubberduck.Settings
//{
//    [XmlType(AnonymousType = true)]
//    public class IndenterSettings : IIndenterSettings
//    {
//        public bool IndentEntireProcedureBody { get; set; }
//        public bool IndentFirstCommentBlock { get; set; }
//        public bool IndentFirstDeclarationBlock { get; set; }
//        public bool AlignCommentsWithCode { get; set; }
//        public bool AlignContinuations { get; set; }
//        public bool IgnoreOperatorsInContinuations { get; set; }
//        public bool IndentCase { get; set; }
//        public bool ForceDebugStatementsInColumn1 { get; set; }
//        public bool ForceCompilerDirectivesInColumn1 { get; set; }
//        public bool IndentCompilerDirectives { get; set; }
//        public bool AlignDims { get; set; }
//        public int AlignDimColumn { get; set; }
//        public bool EnableUndo { get; set; }
//        public EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }
//        public int EndOfLineCommentColumnSpaceAlignment { get; set; }
//        public int IndentSpaces { get; set; }

//        public IndenterSettings()
//        {
//            var tabWidth = 4;
//            var reg = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\6.0\Common", false) ??
//                      Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\7.0\Common", false);
//            if (reg != null)
//            {
//                tabWidth = Convert.ToInt32(reg.GetValue("TabWidth") ?? tabWidth);
//            }

//            IndentEntireProcedureBody = true;
//            IndentFirstCommentBlock = true;
//            IndentFirstDeclarationBlock = true;
//            AlignCommentsWithCode = true;
//            AlignContinuations = true;
//            IgnoreOperatorsInContinuations = true;
//            IndentCase = false;
//            ForceDebugStatementsInColumn1 = false;
//            ForceCompilerDirectivesInColumn1 = false;
//            IndentCompilerDirectives = true;
//            AlignDims = false;
//            AlignDimColumn = 15;
//            EnableUndo = true;
//            EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
//            EndOfLineCommentColumnSpaceAlignment = 50;
//            IndentSpaces = tabWidth;
//        }
//    }
//}