using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class Parameter
    {
        public string Accessibility { get; set; }
        public string Name { get; set; }
        public string AsTypeName { get; set; }
        public string Optional { get; set; }
        public string DefaultValue { get; set; }

        public Parameter()
        {}

        public Parameter(ParameterDeclaration parameter)
        {
            Accessibility = parameter.IsImplicitByRef
                ? string.Empty
                : parameter.IsByRef
                    ? Tokens.ByRef
                    : Tokens.ByVal;

            Name = parameter.IsArray
                ? $"{parameter.IdentifierName}()"
                : parameter.IdentifierName;

            AsTypeName = parameter.AsTypeName;

            Optional = parameter.IsParamArray
                ? Tokens.ParamArray
                : parameter.IsOptional
                    ? Tokens.Optional
                    : string.Empty;

            DefaultValue = parameter.DefaultValue;
        }

        private static string FormatStandardElement(string element) => string.IsNullOrEmpty(element)
            ? string.Empty
            : $"{element} ";

        private string FormattedAsTypeName => string.IsNullOrEmpty(AsTypeName)
            ? string.Empty
            : $"As {AsTypeName} ";

        private string FormattedDefaultValue => string.IsNullOrEmpty(DefaultValue)
            ? string.Empty
            : $"= {DefaultValue}";

        public override string ToString()
        {
            return $"{FormatStandardElement(Optional)}{FormatStandardElement(Accessibility)}{FormatStandardElement(Name)}{FormattedAsTypeName}{FormattedDefaultValue}".Trim();
        }
    }
}
