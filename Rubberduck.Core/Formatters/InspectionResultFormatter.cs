using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Common;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Formatters
{
    public class InspectionResultFormatter : IExportable
    {
        private readonly IInspectionResult _inspectionResult;
        private readonly string _documentName;

        public InspectionResultFormatter(IInspectionResult inspectionResult, string documentName)
        {
            _inspectionResult = inspectionResult;
            _documentName = documentName;
        }

        public object[] ToArray()
        {
            var module = _inspectionResult.QualifiedSelection.QualifiedName;
            return new object[]
            {
                _inspectionResult.Inspection.Severity.ToString(),
                module.ProjectName,
                module.ComponentName,
                _inspectionResult.Description,
                _inspectionResult.QualifiedSelection.Selection.StartLine,
                _inspectionResult.QualifiedSelection.Selection.StartColumn
            };
        }

        public string ToClipboardString()
        {
            var module = _inspectionResult.QualifiedSelection.QualifiedName;
            var documentName = _documentName;

            return string.Format(
                InspectionsUI.QualifiedSelectionInspection,
                _inspectionResult.Inspection.Severity,
                _inspectionResult.Description,
                $"({documentName})",
                module.ProjectName,
                module.ComponentName,
                _inspectionResult.QualifiedSelection.Selection.StartLine);
        }
    }
}
