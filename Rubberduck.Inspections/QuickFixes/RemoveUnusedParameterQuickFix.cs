using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveUnusedParameterQuickFix : IQuickFix
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ParameterNotUsedInspection)
        };

        public RemoveUnusedParameterQuickFix(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            using (var dialog = new RemoveParametersDialog(new RemoveParametersViewModel(_state)))
            {
                var refactoring = new RemoveParametersRefactoring(_vbe,
                    new RemoveParametersPresenterFactory(_vbe, dialog, _state, _messageBox));

                refactoring.QuickFix(_state, result.QualifiedSelection);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveUnusedParameterQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}