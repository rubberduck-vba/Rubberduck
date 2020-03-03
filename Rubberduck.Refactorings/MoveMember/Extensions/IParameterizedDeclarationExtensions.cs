using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember.Extensions
{
    public static class IParameterizedDeclarationExtensions
    {
        public static string BuildFullyDefinedArgumentList(this IParameterizedDeclaration pDeclaration)
        {
            string AsFullyDefinedParameter(ParameterDeclaration p)
            {
                var access = ((VBAParser.ArgContext)p.Context).BYVAL() != null
                    ? Tokens.ByVal
                    : Tokens.ByRef;

                if (p.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false
                    || (p == pDeclaration.Parameters.Last()
                            && (p.ParentDeclaration.DeclarationType.Equals(DeclarationType.PropertyLet)
                                    || p.ParentDeclaration.DeclarationType.Equals(DeclarationType.PropertySet))))
                {
                    access = Tokens.ByVal;
                }

                return $"{access} {p.IdentifierName} {Tokens.As} {p.AsTypeName}";
            }

            var memberParams = pDeclaration.Parameters.ToList()
                .OrderBy(o => o.Selection.StartLine)
                .ThenBy(t => t.Selection.StartColumn)
                .Select(p => AsFullyDefinedParameter(p));

            var improvedArgList = $"({string.Join(", ", memberParams)})";
            return improvedArgList;
        }
    }
}
