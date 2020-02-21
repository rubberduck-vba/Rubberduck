using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public struct PropertyAttributeSet
    {
        public string PropertyName { get; set; }
        public string BackingField { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; }
        public bool GenerateLetter { get; set; }
        public bool GenerateSetter { get; set; }
        public bool UsesSetAssignment { get; set; }
        public bool IsUDTProperty { get; set; }
    }

    public class PropertyGenerator
    {
        public PropertyGenerator() { }

        public string AsPropertyBlock(PropertyAttributeSet attrSet, IIndenter indenter)
        {
            return string.Join(Environment.NewLine, indenter.Indent(AsLines(attrSet), true));
        }

        private string AllPropertyCode(PropertyAttributeSet attrSet) =>
            $"{GetterCode(attrSet)}{(attrSet.GenerateLetter ? LetterCode(attrSet) : string.Empty)}{(attrSet.GenerateSetter ? SetterCode(attrSet) : string.Empty)}";

        private IEnumerable<string> AsLines(PropertyAttributeSet _spec) => AllPropertyCode(_spec).Split(new[] { Environment.NewLine }, StringSplitOptions.None);

        private string GetterCode(PropertyAttributeSet attrSet)
        {
            if (attrSet.GenerateSetter && attrSet.GenerateLetter)
            {
                return string.Join(Environment.NewLine,
                                    $"Public Property Get {attrSet.PropertyName}() As {attrSet.AsTypeName}",
                                    $"    If IsObject({attrSet.BackingField}) Then",
                                    $"        Set {attrSet.PropertyName} = {attrSet.BackingField}",
                                    "    Else",
                                    $"        {attrSet.PropertyName} = {attrSet.BackingField}",
                                    "    End If",
                                    "End Property",
                                    Environment.NewLine);
            }

            return string.Join(Environment.NewLine,
                                $"Public Property Get {attrSet.PropertyName}() As {attrSet.AsTypeName}",
                                $"    {(attrSet.UsesSetAssignment ? "Set " : string.Empty)}{attrSet.PropertyName} = {attrSet.BackingField}",
                                "End Property",
                                Environment.NewLine);
        }

        private string SetterCode(PropertyAttributeSet attrSet)
        {
            if (!attrSet.GenerateSetter)
            {
                return string.Empty;
            }
            return string.Join(Environment.NewLine,
                            $"Public Property Set {attrSet.PropertyName}(ByVal {attrSet.ParameterName} As {attrSet.AsTypeName})",
                            $"    Set {attrSet.BackingField} = {attrSet.ParameterName}",
                            "End Property",
                            Environment.NewLine);
        }

        private string LetterCode(PropertyAttributeSet attrSet)
        {
            if (!attrSet.GenerateLetter)
            {
                return string.Empty;
            }

            var byVal_byRef = attrSet.IsUDTProperty ? Tokens.ByRef : Tokens.ByVal;

            return string.Join(Environment.NewLine,
                                $"Public Property Let {attrSet.PropertyName}({byVal_byRef} {attrSet.ParameterName} As {attrSet.AsTypeName})",
                                $"    {attrSet.BackingField} = {attrSet.ParameterName}",
                                "End Property",
                                Environment.NewLine);
        }
    }
}
