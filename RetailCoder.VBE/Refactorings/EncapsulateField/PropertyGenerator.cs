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

        public string AllPropertyCode
        {
            get
            {
                return GetterCode +
                        (GenerateLetter ? LetterCode : string.Empty) +
                        (GenerateSetter ? SetterCode : string.Empty);
            }
        }

        public string GetterCode
        {
            get
            {
                if (GenerateSetter && GenerateLetter)
                {
                    return string.Join(Environment.NewLine,
                        string.Format("Public Property Get {0}() As {1}", PropertyName, AsTypeName),
                        string.Format("    If IsObject({0}) Then", BackingField),
                        string.Format("        Set {0} = {1}", PropertyName, BackingField),
                                      "    Else",
                        string.Format("        {0} = {1}", PropertyName, BackingField),
                                      "    End If",
                                      "End Property",
                                      Environment.NewLine);
                }

                return string.Join(Environment.NewLine,
                    string.Format("Public Property Get {0}() As {1}", PropertyName, AsTypeName),
                    string.Format("    {0}{1} = {2}", GenerateSetter ? "Set " : string.Empty, PropertyName, BackingField),
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
                    string.Format("Public Property Set {0}(ByVal {1} As {2})", PropertyName, ParameterName, AsTypeName),
                    string.Format("    Set {0} = {1}", BackingField, ParameterName),
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
                    string.Format("Public Property Let {0}(ByVal {1} As {2})", PropertyName, ParameterName, AsTypeName),
                    string.Format("    {0} = {1}", BackingField, ParameterName),
                                  "End Property",
                                  Environment.NewLine);
            }
        }
    }
}
