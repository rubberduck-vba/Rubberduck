using System;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IntroduceLocalVariableQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public IntroduceLocalVariableQuickFix(RubberduckParserState state)
            : base(typeof(UndeclaredVariableInspection))
        {
            _state = state;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        public override void Fix(IInspectionResult result)
        {
            var instruction = $"{Environment.NewLine}Dim {result.Target.IdentifierName} As Variant{Environment.NewLine}";
            _state.GetRewriter(result.Target).InsertBefore(result.Target.Context.Start.TokenIndex, instruction);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.IntroduceLocalVariableQuickFix;
    }
}