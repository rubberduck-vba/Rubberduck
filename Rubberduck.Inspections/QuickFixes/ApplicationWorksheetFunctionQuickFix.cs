using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ApplicationWorksheetFunctionQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> {typeof(ApplicationWorksheetFunctionInspection) };
        
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
            
            var oldContent = module.GetLines(result.QualifiedSelection.Selection);
            var newCall = $"WorksheetFunction.{result.Target.IdentifierName}";

            var start = result.QualifiedSelection.Selection.StartColumn - 1;
            //The member being called will always be a single token, so this will always be safe (it will be a single line).
            var end = result.QualifiedSelection.Selection.EndColumn - 1;

            var newContent = oldContent.Substring(0, start) + newCall + 
                (oldContent.Length > end
                ? oldContent.Substring(end, oldContent.Length - end)
                : string.Empty);

            module.DeleteLines(result.QualifiedSelection.Selection);
            module.InsertLines(result.QualifiedSelection.Selection.StartLine, newContent);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ApplicationWorksheetFunctionQuickFix;
        }

        public bool CanFixInProcedure { get; } = true;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}
