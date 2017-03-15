using System;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class PropertyGenerator
    {
        public string PropertyName { get; set; }
        public string BackingField { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; }
        public bool GenerateLetter { get; set; }
        public bool GenerateSetter { get; set; }

        public string AllPropertyCode => GetterCode +
                                         (GenerateLetter ? LetterCode : string.Empty) +
                                         (GenerateSetter ? SetterCode : string.Empty);

        public string GetterCode
        {
            get
            {
                if (GenerateSetter && GenerateLetter)
                {
                    return string.Join(Environment.NewLine,
                                       $"Public Property Get {PropertyName}() As {AsTypeName}",
                                       $"    If IsObject({BackingField}) Then",
                                       $"        Set {PropertyName} = {BackingField}",
                                       "    Else",
                                       $"        {PropertyName} = {BackingField}",
                                       "    End If",
                                       "End Property",
                                       Environment.NewLine);
                }

                return string.Join(Environment.NewLine,
                                   $"Public Property Get {PropertyName}() As {AsTypeName}",
                                   $"    {(GenerateSetter ? "Set " : string.Empty)}{PropertyName} = {BackingField}",
                                   "End Property",
                                   Environment.NewLine);
            }
        }

        public string SetterCode
        {
            get
            {
                if (!GenerateSetter)
                {
                    return string.Empty;
                }
                return string.Join(Environment.NewLine,
                                   $"Public Property Set {PropertyName}(ByVal {ParameterName} As {AsTypeName})",
                                   $"    Set {BackingField} = {ParameterName}",
                                   "End Property",
                                   Environment.NewLine);
            }
        }

        public string LetterCode
        {
            get
            {
                if (!GenerateLetter)
                {
                    return string.Empty;
                }
                return string.Join(Environment.NewLine,
                                   $"Public Property Let {PropertyName}(ByVal {ParameterName} As {AsTypeName})",
                                   $"    {BackingField} = {ParameterName}",
                                   "End Property",
                                   Environment.NewLine);
            }
        }
    }
}
