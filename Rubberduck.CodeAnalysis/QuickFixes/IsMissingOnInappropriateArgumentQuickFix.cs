using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class IsMissingOnInappropriateArgumentQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public IsMissingOnInappropriateArgumentQuickFix(RubberduckParserState state)
            : base(typeof(IsMissingOnInappropriateArgumentInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            if (!(result.Properties is ParameterDeclaration parameter))
            {
                Logger.Trace(
                    $"Properties for IsMissingOnInappropriateArgumentQuickFix was {(result.Properties == null ? "null" : "not a ParameterDeclaration")}.");
                return;
            }

            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            if (!result.Context.TryGetAncestor<VBAParser.LExprContext>(out var context))
            {
                Logger.Trace("IsMissingOnInappropriateArgumentQuickFix could not locate containing LExprContext for replacement.");
                return;
            }

            if (parameter.IsParamArray || parameter.IsArray)
            {
                rewriter.Replace(context, $"{Tokens.LBound}({parameter.IdentifierName}) > {Tokens.UBound}({parameter.IdentifierName})");
                return;
            }

            if (!string.IsNullOrEmpty(parameter.DefaultValue))
            {
                if (parameter.DefaultValue.Equals("\"\""))
                {
                    rewriter.Replace(context, $"{parameter.IdentifierName} = {Tokens.vbNullString}");
                }
                else if (parameter.DefaultValue.Equals(Tokens.Nothing, StringComparison.InvariantCultureIgnoreCase))
                {
                    rewriter.Replace(context, $"{parameter.IdentifierName} Is {Tokens.Nothing}");
                }
                else
                {
                    rewriter.Replace(context, $"{parameter.IdentifierName} = {parameter.DefaultValue}");
                }
                return;
            }
            rewriter.Replace(context, UninitializedComparisonForParameter(parameter));
        }

        private static readonly Dictionary<string, string> BaseTypeUninitializedValues = new Dictionary<string, string>
        {
            { Tokens.Boolean.ToUpper(), Tokens.False },
            { Tokens.Byte.ToUpper(), "0" },
            { Tokens.Currency.ToUpper(), "0" },
            { Tokens.Date.ToUpper(), "CDate(0)" },
            { Tokens.Decimal.ToUpper(), "0" },
            { Tokens.Double.ToUpper(), "0" },
            { Tokens.Integer.ToUpper(), "0" },
            { Tokens.Long.ToUpper(), "0" },
            { Tokens.LongLong.ToUpper(), "0" },
            { Tokens.LongPtr.ToUpper(),  "0"  },
            { Tokens.Single.ToUpper(), "0" },
            { Tokens.String.ToUpper(), Tokens.vbNullString }
        };

        private string UninitializedComparisonForParameter(ParameterDeclaration parameter)
        {
            var type = parameter.AsTypeName?.ToUpper() ?? string.Empty;
            if (string.IsNullOrEmpty(type))
            {
                type = parameter.HasTypeHint
                    ? SymbolList.TypeHintToTypeName[parameter.TypeHint].ToUpper()
                    : Tokens.Variant.ToUpper();
            }

            if (BaseTypeUninitializedValues.ContainsKey(type))
            {
                return $"{parameter.IdentifierName} = {BaseTypeUninitializedValues[type]}";
            }

            if (type.Equals(Tokens.Object, StringComparison.InvariantCultureIgnoreCase))
            {
                return $"{parameter.IdentifierName} Is {Tokens.Nothing}";
            }

            if (type.Equals(Tokens.Object, StringComparison.InvariantCultureIgnoreCase) || parameter.AsTypeDeclaration == null)
            {
                return $"IsEmpty({parameter.IdentifierName})";
            }

            switch (parameter.AsTypeDeclaration.DeclarationType)
            {
                case DeclarationType.ClassModule:
                    return $"{parameter.IdentifierName} Is {Tokens.Nothing}";
                case DeclarationType.Enumeration:
                    var members = _state.DeclarationFinder.AllDeclarations.OfType<ValuedDeclaration>()
                        .FirstOrDefault(decl =>
                            ReferenceEquals(decl.ParentDeclaration, parameter.AsTypeDeclaration) &&
                            decl.Expression.Equals("0"));
                    return $"{parameter.IdentifierName} = {members?.IdentifierName ?? "0"}";
                default:
                    return $"IsError({parameter.IdentifierName})";
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IsMissingOnInappropriateArgumentQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}
