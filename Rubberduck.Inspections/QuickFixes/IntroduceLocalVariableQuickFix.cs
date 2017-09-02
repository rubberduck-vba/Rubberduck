using System;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IntroduceLocalVariableQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public IntroduceLocalVariableQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<UndeclaredVariableInspection>());
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;

        public void Fix(IInspectionResult result)
        {
            var instruction = $"{Environment.NewLine}Dim {result.Target.IdentifierName} As Variant{Environment.NewLine}";
            _state.GetRewriter(result.Target).InsertBefore(result.Target.Context.Start.TokenIndex, instruction);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.IntroduceLocalVariableQuickFix;
        }
    }
}