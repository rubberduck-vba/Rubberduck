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
        public virtual bool IndentEnumTypeAsProcedure { get; set; }
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
        public virtual EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }
        public virtual int EndOfLineCommentColumnSpaceAlignment { get; set; }
        public virtual int IndentSpaces { get; set; }

        public IndenterSettings()
        {
            var tabWidth = 4;
            try
            {
                var reg = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\6.0\Common", false) ??
                          Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\7.0\Common", false);
                if (reg != null)
                {
                    tabWidth = Convert.ToInt32(reg.GetValue("TabWidth") ?? tabWidth);
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }

            // Mocking requires these to be virtual.
            // ReSharper disable DoNotCallOverridableMethodsInConstructor
            IndentEntireProcedureBody = true;
            IndentFirstCommentBlock = true;
            IndentEnumTypeAsProcedure = false;
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
            EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
            EndOfLineCommentColumnSpaceAlignment = 50;
            IndentSpaces = tabWidth;
            // ReSharper restore DoNotCallOverridableMethodsInConstructor
        }

        private const string LegacySettingsSubKey = @"Software\VB and VBA Program Settings\Office Automation Ltd.\Smart Indenter";
        public bool LegacySettingsExist()
        {
            try
            {
                return (Registry.CurrentUser.OpenSubKey(LegacySettingsSubKey, false) != null);
            }
            catch 
            {
                return false;
            }
            
        }

        public void LoadLegacyFromRegistry()
        {
            try
            {
                var reg = Registry.CurrentUser.OpenSubKey(LegacySettingsSubKey, false);
                if (reg == null) return;
                IndentEntireProcedureBody = GetSmartIndenterBoolean(reg, "IndentProc", IndentEntireProcedureBody);
                IndentFirstCommentBlock = GetSmartIndenterBoolean(reg, "IndentFirst", IndentFirstCommentBlock);
                IndentFirstDeclarationBlock = GetSmartIndenterBoolean(reg, "IndentDim", IndentFirstDeclarationBlock);
                AlignCommentsWithCode = GetSmartIndenterBoolean(reg, "IndentCmt", AlignCommentsWithCode);
                AlignContinuations = GetSmartIndenterBoolean(reg, "AlignContinued", AlignContinuations);
                IgnoreOperatorsInContinuations = GetSmartIndenterBoolean(reg, "AlignIgnoreOps",
                    IgnoreOperatorsInContinuations);
                IndentCase = GetSmartIndenterBoolean(reg, "IndentCase", IndentCase);
                ForceDebugStatementsInColumn1 = GetSmartIndenterBoolean(reg, "DebugCol1", ForceDebugStatementsInColumn1);
                ForceCompilerDirectivesInColumn1 = GetSmartIndenterBoolean(reg, "CompilerCol1",
                    ForceCompilerDirectivesInColumn1);
                IndentCompilerDirectives = GetSmartIndenterBoolean(reg, "IndentCompiler", IndentCompilerDirectives);
                AlignDims = GetSmartIndenterBoolean(reg, "AlignDim", AlignDims);
                AlignDimColumn = Convert.ToInt32(reg.GetValue("AlignDimCol") ?? AlignDimColumn);

                var eolSytle = reg.GetValue("EOLComments") as string;
                if (!string.IsNullOrEmpty(eolSytle))
                {
                    switch (eolSytle)
                    {
                        case "Absolute":
                            EndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
                            break;
                        case "SameGap":
                            EndOfLineCommentStyle = EndOfLineCommentStyle.SameGap;
                            break;
                        case "StandardGap":
                            EndOfLineCommentStyle = EndOfLineCommentStyle.StandardGap;
                            break;
                        case "AlignInCol":
                            EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
                            break;
                    }
                }
                EndOfLineCommentColumnSpaceAlignment =
                    Convert.ToInt32(reg.GetValue("EOLAlignCol") ?? EndOfLineCommentColumnSpaceAlignment);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }
        }

        private static bool GetSmartIndenterBoolean(RegistryKey key, string name, bool current)
        {
            var value = key.GetValue(name) as string;
            return string.IsNullOrEmpty(value) ? current : value.Trim().Equals("Y");
        }
    }
}
