using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.Properties;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class CodeInspectionResultGridViewItem
    {
        public CodeInspectionResultGridViewItem(CodeInspectionResultBase result)
        {
            _item = result;
            _severity = GetSeverityIcon(result.Severity);
            _project = result.Node.Instruction.Line.ProjectName;
            _component = result.Node.Instruction.Line.ComponentName;
            _selection = result.Node.Instruction.Selection;
            _issue = result.Name;
            _quickFix = FirstOrDefaultQuickFix(result.GetQuickFixes());
        }

        private readonly CodeInspectionResultBase _item;
        public CodeInspectionResultBase GetInspectionResultItem()
        {
            return _item;
        }

        private object _quickFix;
        private Action<VBE> FirstOrDefaultQuickFix(IDictionary<string, Action<VBE>> fixes)
        {
            return fixes.FirstOrDefault().Value;
        }

        private static readonly IDictionary<CodeInspectionSeverity, Bitmap> _severityIcons =
            new Dictionary<CodeInspectionSeverity, Bitmap>
            {
                { CodeInspectionSeverity.DoNotShow, null },
                { CodeInspectionSeverity.Hint, Resources.Information },
                { CodeInspectionSeverity.Suggestion, Resources.Alert },
                { CodeInspectionSeverity.Warning, Resources.Warning },
                { CodeInspectionSeverity.Error, Resources.Critical }
            };

        private Image GetSeverityIcon(CodeInspectionSeverity severity)
        {
            var image = _severityIcons[severity];
            image.MakeTransparent(Color.Fuchsia);
            return image;
        }

        private readonly Image _severity;
        public Image Severity
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

        private readonly Selection _selection;
        public int Line
        {
            get { return _selection.StartLine; }
        }

        private readonly string _issue;
        public string Issue
        {
            get { return _issue; }
        }
    }
}