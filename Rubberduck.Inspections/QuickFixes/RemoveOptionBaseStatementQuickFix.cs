using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveOptionBaseStatementQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(OptionBaseZeroInspection)
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
            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var lines = module.GetLines(result.QualifiedSelection.Selection).Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var newContent = result.QualifiedSelection.Selection.LineCount != 1
                ? lines[0].Remove(result.QualifiedSelection.Selection.StartColumn - 1)
                : lines[0].Remove(result.QualifiedSelection.Selection.StartColumn - 1, result.QualifiedSelection.Selection.EndColumn - result.QualifiedSelection.Selection.StartColumn);
            
            if (result.QualifiedSelection.Selection.LineCount != 1)
            {
                newContent += lines.Last().Remove(0, result.QualifiedSelection.Selection.EndColumn - 1);
            }

            module.DeleteLines(result.QualifiedSelection.Selection);
            module.InsertLines(result.QualifiedSelection.Selection.StartLine, newContent);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveOptionBaseStatementQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => false;
        public bool CanFixInProject => false;
    }
}