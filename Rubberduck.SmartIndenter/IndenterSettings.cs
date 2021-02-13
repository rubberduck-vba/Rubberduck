using System;
using System.Xml.Serialization;
using Microsoft.Win32;

namespace Rubberduck.SmartIndenter
{
    [XmlType(AnonymousType = true)]
    public class IndenterSettings : IIndenterSettings, IEquatable<IndenterSettings>
    {
        // These have to be int to allow the settings UI to bind them.
        public const int MaximumAlignDimColumn = 100;
        public const int MaximumEndOfLineCommentColumnSpaceAlignment = 100;
        public const int MaximumIndentSpaces = 32;
        public const int MaximumVerticalSpacing = 2;     

        public virtual bool IndentEntireProcedureBody { get; set; }
        public virtual bool IndentEnumTypeAsProcedure { get; set; }
        public virtual bool IndentFirstCommentBlock { get; set; }       
        public virtual bool IndentFirstDeclarationBlock { get; set; }
        public virtual bool IgnoreEmptyLinesInFirstBlocks { get; set; }
        public virtual bool AlignCommentsWithCode { get; set; }
        public virtual bool AlignContinuations { get; set; }
        public virtual bool IgnoreOperatorsInContinuations { get; set; }
        public virtual bool IndentCase { get; set; }

        private bool _forceDebugs;
        public virtual bool ForceDebugStatementsInColumn1
        {
            get => _forceDebugs;
            set
            {
                _forceDebugs = value;
                _forceDebugPrint = _forceDebugs;
                _forceDebugAssert = _forceDebugs;
                _forceStop = _forceDebugs;
            }
        }

        private bool _forceDebugPrint;
        public virtual bool ForceDebugPrintInColumn1
        {
            get => _forceDebugPrint;
            set
            {
                _forceDebugPrint = value;
                if (!_forceDebugPrint && !_forceDebugAssert && !_forceStop)
                {
                    _forceDebugs = false;
                }
            }
        }

        private bool _forceDebugAssert;
        public virtual bool ForceDebugAssertInColumn1
        {
            get => _forceDebugAssert;
            set
            {
                _forceDebugAssert = value;
                if (!_forceDebugPrint && !_forceDebugAssert && !_forceStop)
                {
                    _forceDebugs = false;
                }
            }
        }

        private bool _forceStop;
        public virtual bool ForceStopInColumn1
    {
            get => _forceStop;
            set
            {
                _forceStop = value;
                if (!_forceDebugPrint && !_forceDebugAssert && !_forceStop)
                {
                    _forceDebugs = false;
                }
            }
        }

        public virtual bool ForceCompilerDirectivesInColumn1 { get; set; }
        public virtual bool IndentCompilerDirectives { get; set; }
        public virtual bool AlignDims { get; set; }

        private int _dimAlignment;
        public virtual int AlignDimColumn
        {
            get => _dimAlignment;
            set => _dimAlignment = value > MaximumAlignDimColumn ? MaximumAlignDimColumn : Math.Max(value, 0);
        }

        public virtual EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }

        public virtual EmptyLineHandling EmptyLineHandlingMethod { get; set; }

        private int _commentAlignment;
        public virtual int EndOfLineCommentColumnSpaceAlignment
        {
            get => _commentAlignment;
            set => _commentAlignment = value > MaximumEndOfLineCommentColumnSpaceAlignment
                ? MaximumEndOfLineCommentColumnSpaceAlignment
                : value;
        }

        private int _indentSpaces;
        public virtual int IndentSpaces
        {
            get => _indentSpaces;
            set => _indentSpaces = value > MaximumIndentSpaces ? MaximumIndentSpaces : Math.Max(value, 0);
        }

        public virtual bool VerticallySpaceProcedures { get; set; }

        private int _procedureSpacing;
        public virtual int LinesBetweenProcedures
        {
            get => _procedureSpacing;
            set => _procedureSpacing = value > MaximumVerticalSpacing ? MaximumVerticalSpacing : Math.Max(value, 0);
        }

        public virtual bool GroupRelatedProperties { get; set; }

        /// <summary>
        /// Use this ctor for unit testing.
        /// </summary>
        public IndenterSettings() :this(true) { }

        /// <summary>
        /// Creates an IndenterSettings. 
        /// </summary>
        /// <param name="skipRegistry">If false, the ctor will attempt to load the current tab width from the registry.</param>
        public IndenterSettings(bool skipRegistry)
        {
            var tabWidth = 4;
            if (!skipRegistry)
            {
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
                catch
                { }
            }

            // Mocking requires these to be virtual.
            // ReSharper disable DoNotCallOverridableMethodsInConstructor
            IndentEntireProcedureBody = true;
            IndentEnumTypeAsProcedure = false;
            IndentFirstCommentBlock = true;            
            IndentFirstDeclarationBlock = true;
            IgnoreEmptyLinesInFirstBlocks = false;
            AlignCommentsWithCode = true;
            AlignContinuations = true;
            IgnoreOperatorsInContinuations = true;
            IndentCase = false;
            ForceDebugStatementsInColumn1 = false;
            ForceDebugPrintInColumn1 = false;
            ForceDebugAssertInColumn1 = false;
            ForceStopInColumn1 = false;
            ForceCompilerDirectivesInColumn1 = false;
            IndentCompilerDirectives = true;
            AlignDims = false;
            AlignDimColumn = 15;
            EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
            EmptyLineHandlingMethod = EmptyLineHandling.Ignore;
            EndOfLineCommentColumnSpaceAlignment = 50;
            IndentSpaces = tabWidth;
            VerticallySpaceProcedures = true;
            LinesBetweenProcedures = 1;
            // ReSharper restore DoNotCallOverridableMethodsInConstructor
        }

        public bool Equals(IndenterSettings other)
        {
            return other != null &&
                   IndentEntireProcedureBody == other.IndentEntireProcedureBody &&
                   IndentFirstCommentBlock == other.IndentFirstCommentBlock &&
                   IndentEnumTypeAsProcedure == other.IndentEnumTypeAsProcedure &&
                   IndentFirstDeclarationBlock == other.IndentFirstDeclarationBlock &&
                   IgnoreEmptyLinesInFirstBlocks == other.IgnoreEmptyLinesInFirstBlocks &&
                   AlignCommentsWithCode == other.AlignCommentsWithCode &&
                   AlignContinuations == other.AlignContinuations &&
                   IgnoreOperatorsInContinuations == other.IgnoreOperatorsInContinuations &&
                   IndentCase == other.IndentCase &&
                   ForceDebugStatementsInColumn1 == other.ForceDebugStatementsInColumn1 &&
                   ForceDebugPrintInColumn1 == other.ForceDebugPrintInColumn1 &&
                   ForceDebugAssertInColumn1 == other.ForceDebugAssertInColumn1 &&
                   ForceStopInColumn1 == other.ForceStopInColumn1 &&
                   ForceCompilerDirectivesInColumn1 == other.ForceCompilerDirectivesInColumn1 &&
                   IndentCompilerDirectives == other.IndentCompilerDirectives &&
                   AlignDims == other.AlignDims &&
                   AlignDimColumn == other.AlignDimColumn &&
                   EndOfLineCommentStyle == other.EndOfLineCommentStyle &&
                   EndOfLineCommentColumnSpaceAlignment == other.EndOfLineCommentColumnSpaceAlignment &&
                   IndentSpaces == other.IndentSpaces &&
                   VerticallySpaceProcedures == other.VerticallySpaceProcedures &&
                   LinesBetweenProcedures == other.LinesBetweenProcedures &&
                   EmptyLineHandlingMethod == other.EmptyLineHandlingMethod;
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
                IgnoreOperatorsInContinuations = GetSmartIndenterBoolean(reg, "AlignIgnoreOps", IgnoreOperatorsInContinuations);
                IndentCase = GetSmartIndenterBoolean(reg, "IndentCase", IndentCase);
                ForceDebugStatementsInColumn1 = GetSmartIndenterBoolean(reg, "DebugCol1", ForceDebugStatementsInColumn1);
                ForceDebugPrintInColumn1 = ForceDebugStatementsInColumn1;
                ForceDebugAssertInColumn1 = ForceDebugStatementsInColumn1;
                ForceStopInColumn1 = ForceDebugStatementsInColumn1;
                ForceCompilerDirectivesInColumn1 = GetSmartIndenterBoolean(reg, "CompilerCol1", ForceCompilerDirectivesInColumn1);
                IndentCompilerDirectives = GetSmartIndenterBoolean(reg, "IndentCompiler", IndentCompilerDirectives);
                AlignDims = GetSmartIndenterBoolean(reg, "AlignDim", AlignDims);
                AlignDimColumn = GetSmartIndenterNumeric(reg, "AlignDimCol", AlignDimColumn, MaximumAlignDimColumn);

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
                EndOfLineCommentColumnSpaceAlignment = GetSmartIndenterNumeric(reg, "EOLAlignCol",
                    EndOfLineCommentColumnSpaceAlignment, MaximumEndOfLineCommentColumnSpaceAlignment);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }
        }

        private static bool GetSmartIndenterBoolean(RegistryKey key, string name, bool current)
        {
            var value = key.GetValue(name) as string;
            return string.IsNullOrEmpty(value) ? current : value.Trim().Equals("Y");
        }

        private static int GetSmartIndenterNumeric(RegistryKey key, string name, int current, int max)
        {
            try
            {
                var value = (int)key.GetValue(name);
                return value < 0 ? current : Math.Min(value, max);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }
            return current;
        }
    }
}
