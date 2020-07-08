using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class DeclarationToMemberSignatureConverter : IValueConverter
    {
        public class Parameter
        {
            public string ParamAccessibility { get; set; }
            public string ParamName { get; set; }
            public string ParamType { get; set; }

            public override string ToString()
            {
                return $"{ParamAccessibility} {ParamName} {Tokens.As} {ParamType}";
            }
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var typedValue = (Declaration)value;
            return FullMemberSignature(typedValue);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        private string FullMemberSignature(Declaration member)
        {
            var signature = $"{GetMethodType(member)} {member.IdentifierName}({string.Join(", ", GetMemberParameters(member))})";

            return member.AsTypeName == null ? signature : $"{signature} {Tokens.As} {member.AsTypeName}";
        }

        private List<Parameter> GetMemberParameters(Declaration member)
        {
            var parameters = new List<Parameter>();

            if (member is IParameterizedDeclaration memberWithParams)
            {
                parameters = memberWithParams.Parameters
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .Select(p => new Parameter
                    {
                        ParamAccessibility =
                            ((VBAParser.ArgContext)p.Context).BYVAL() != null ? Tokens.ByVal : Tokens.ByRef,
                        ParamName = p.IdentifierName,
                        ParamType = p.AsTypeName
                    })
                    .ToList();
            }
            else
            {
                return new List<Parameter>();
            }

            if (GetMethodType(member) == $"{Tokens.Property} {Tokens.Get}")
            {
                parameters = parameters.Take(parameters.Count - 1).ToList();
            }

            return parameters;
        }

        private string GetMethodType(Declaration member)
        {
            var context = member.Context;

            if (context is VBAParser.SubStmtContext)
            {
                return Tokens.Sub;
            }

            if (context is VBAParser.FunctionStmtContext)
            {
                return Tokens.Function;
            }

            if (context is VBAParser.PropertyGetStmtContext)
            {
                return $"{Tokens.Property} {Tokens.Get}";
            }

            if (context is VBAParser.PropertyLetStmtContext)
            {
                return $"{Tokens.Property} {Tokens.Let}";
            }

            if (context is VBAParser.PropertySetStmtContext)
            {
                return $"{Tokens.Property} {Tokens.Set}";
            }

            throw new ArgumentException(nameof(member));
        }
    }
}
