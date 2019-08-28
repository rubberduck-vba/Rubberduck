using Rubberduck.Common;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Inspections.Abstract;
using System.IO;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Formatters
{
    public class InspectionResultFormatter : IExportable
    {
        private readonly IInspectionResult _inspectionResult;

        public InspectionResultFormatter(IInspectionResult inspectionResult)
        {
            _inspectionResult = inspectionResult;
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

        /// <summary>
        /// WARNING: This property can have side effects. It can change the ActiveVBProject if the result has a null Declaration, 
        /// which causes a flicker in the VBE. This should only be called if it is *absolutely* necessary.
        /// </summary>
        public string ToClipboardString()
        {
            var module = _inspectionResult.QualifiedSelection.QualifiedName;
            var documentName = _inspectionResult.Target != null
                ? _inspectionResult.Target.ProjectDisplayName
                : string.Empty;

            //todo: Find a sane way to reimplement this.
            //if (string.IsNullOrEmpty(documentName))
            //{
            //    var component = module.Component;
            //    documentName = component != null 
            //        ? component.ParentProject.ProjectDisplayName 
            //        : string.Empty;
            //}

            if (string.IsNullOrEmpty(documentName))
            {
                documentName = Path.GetFileName(module.ProjectPath);
            }

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
