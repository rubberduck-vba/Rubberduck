using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NLog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
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
            if (!(result.Target is ParameterDeclaration parameter))
            {
                Logger.Trace(
                    $"Target for IsMissingOnInappropriateArgumentQuickFix was {(result.Target == null ? "null" : "not a ParameterDeclaration")}.");
                return;
            }

            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            if (parameter.IsParamArray)
            {
                rewriter.Replace(result.Context, $"{Tokens.LBound}({parameter.IdentifierName}) > {Tokens.UBound}({parameter.IdentifierName})");
            }
            else if (!string.IsNullOrEmpty(parameter.DefaultValue))
            {
                if (parameter.DefaultValue.Equals("\"\""))
                {
                    rewriter.Replace(result.Context, $"{parameter.IdentifierName} = {Tokens.vbNullString}");
                }

                if (parameter.DefaultValue.Equals(Tokens.Nothing))
                {
                    rewriter.Replace(result.Context, $"{parameter.IdentifierName} Is {Tokens.Nothing}");
                }

                rewriter.Replace(result.Context, $"{parameter.IdentifierName} = {parameter.DefaultValue}");
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IsMissingOnInappropriateArgumentQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
