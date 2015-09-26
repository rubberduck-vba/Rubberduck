using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.Properties;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionResultGridViewItem
    {
        public CodeInspectionResultGridViewItem(ICodeInspectionResult result)
        {
            _item = result;
            _icon = GetSeverityIcon(result.Inspection.Severity);
            _severity = result.Inspection.Severity;
            _selection = result.QualifiedSelection;
            _issue = result.Name;
            _quickFix = result.QuickFixes.FirstOrDefault();

            _project = _selection.QualifiedName.ProjectName;
            _component = _selection.QualifiedName.ComponentName;
        }

        private readonly ICodeInspectionResult _item;
        public ICodeInspectionResult GetInspectionResultItem()
        {
            return _item;
        }

        private object _quickFix;
        private static readonly IDictionary<CodeInspectionSeverity, Bitmap> _severityIcons =
            new Dictionary<CodeInspectionSeverity, Bitmap>
            {
                { CodeInspectionSeverity.DoNotShow, null },
                { CodeInspectionSeverity.Hint, Resources.information_white },
                { CodeInspectionSeverity.Suggestion, Resources.information },
                { CodeInspectionSeverity.Warning, Resources.exclamation },
                { CodeInspectionSeverity.Error, Resources.cross_circle }
            };

        private Image GetSeverityIcon(CodeInspectionSeverity severity)
        {
            var image = _severityIcons[severity];
            image.MakeTransparent(Color.Fuchsia);
            return image;
        }

        private readonly Image _icon;
        public Image Icon
        {
            get { return _icon; }
        }

        private readonly CodeInspectionSeverity _severity;
        public CodeInspectionSeverity Severity
        {
            get { return _severity; }
        }

        private readonly string _project;
        public string Project
        {
            get { return _project; }
        }

        private readonly string _component;
        public string Component
        {
            get { return _component; }
        }

        private readonly QualifiedSelection _selection;
        public int Line
        {
            get { return _selection.Selection.StartLine; }
        }

        private readonly string _issue;
        public string Issue
        {
            get { return _issue; }
        }
    }
}