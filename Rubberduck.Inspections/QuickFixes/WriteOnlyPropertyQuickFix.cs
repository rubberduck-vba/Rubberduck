using System;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class WriteOnlyPropertyQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public WriteOnlyPropertyQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<WriteOnlyPropertyInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var parameters = ((IParameterizedDeclaration) result.Target).Parameters.Cast<ParameterDeclaration>().ToList();

            var signatureParams = parameters.Except(new[] {parameters.Last()}).Select(GetParamText);
            var propertyGet = "Public Property Get " + result.Target.IdentifierName + "(" + string.Join(", ", signatureParams) + ") As " +
                parameters.Last().AsTypeName + Environment.NewLine + "End Property" + Environment.NewLine + Environment.NewLine;

            var rewriter = _state.GetRewriter(result.Target);
            rewriter.InsertBefore(result.Target.Context.Start.TokenIndex, propertyGet);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.WriteOnlyPropertyQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;

        private string GetParamText(ParameterDeclaration param)
        {
            return (((VBAParser.ArgContext)param.Context).BYVAL() == null ? "ByRef " : "ByVal ") + param.IdentifierName + " As " + param.AsTypeName;
        }
    }
}
