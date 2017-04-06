using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class MakeSingleLineParameterQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(MultilineParameterInspection)
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
            var selection = result.QualifiedSelection.Selection;

            var lines = module.GetLines(selection.StartLine, selection.EndLine - selection.StartLine + 1);

            var startLine = module.GetLines(selection.StartLine, 1);
            var endLine = module.GetLines(selection.EndLine, 1);

            var adjustedStartColumn = selection.StartColumn - 1;
            var adjustedEndColumn = lines.Length - (endLine.Length - (selection.EndColumn > endLine.Length ? endLine.Length : selection.EndColumn - 1));

            var parameter = lines.Substring(adjustedStartColumn,
                adjustedEndColumn - adjustedStartColumn)
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            var start = startLine.Remove(adjustedStartColumn);
            var end = lines.Remove(0, adjustedEndColumn);

            module.ReplaceLine(selection.StartLine, start + parameter + end);
            module.DeleteLines(selection.StartLine + 1, selection.EndLine - selection.StartLine);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.MakeSingleLineParameterQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}
