using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IntroduceLocalVariableQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(UndeclaredVariableInspection)
        };

        public IntroduceLocalVariableQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

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