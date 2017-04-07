using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SpecifyExplicitPublicModifierQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ImplicitPublicMemberInspection)
        };

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
            var selection = result.Context.GetSelection();
            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var signatureLine = selection.StartLine;

            var oldContent = module.GetLines(signatureLine, 1);
            var newContent = Tokens.Public + ' ' + oldContent;

            module.DeleteLines(signatureLine);
            module.InsertLines(signatureLine, newContent);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SpecifyExplicitPublicModifierQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}