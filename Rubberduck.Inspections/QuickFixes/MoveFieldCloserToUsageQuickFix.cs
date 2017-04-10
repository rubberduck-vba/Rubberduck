using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MoveFieldCloserToUsageQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(MoveFieldCloserToUsageInspection)
        };

        public MoveFieldCloserToUsageQuickFix(RubberduckParserState state, IMessageBox messageBox)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var vbe = result.Target.Project.VBE;

            var refactoring = new MoveCloserToUsageRefactoring(vbe, _state, _messageBox);
            refactoring.Refactor(result.Target);
        }

        public string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, result.Target.IdentifierName);
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}