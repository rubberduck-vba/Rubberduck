using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.Properties;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class CodeInspectionsToolbar
    {
        private readonly VBE _vbe;
        private readonly IEnumerable<IInspection> _inspections;
        private readonly IRubberduckParser _parser;

        private CodeInspectionResultBase[] _issues;
        private int _currentIssue;

        public CodeInspectionsToolbar(VBE vbe, IRubberduckParser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _parser = parser;
            _inspections = inspections;
        }

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;
        private CommandBarButton _quickFixButton;
        private CommandBarButton _navigatePreviousButton;
        private CommandBarButton _navigateNextButton;

        public void Initialize()
        {
            var toolbar = _vbe.CommandBars.Add("Code Inspections", Temporary: true);
            _refreshButton = (CommandBarButton)toolbar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            _refreshButton.TooltipText = "Run code inspections";

            var refreshIcon = Resources.Refresh;
            refreshIcon.MakeTransparent(Color.Magenta);
            Menu.SetButtonImage(_refreshButton, refreshIcon);

            _statusButton = (CommandBarButton)toolbar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            _statusButton.Caption = "0 issues";
            _statusButton.FaceId = 463; // Resources.Warning doesn't look good here
            _statusButton.Style = MsoButtonStyle.msoButtonIconAndCaption;

            _quickFixButton = (CommandBarButton)toolbar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            _quickFixButton.Caption = "Fix";
            _quickFixButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _quickFixButton.FaceId = 305; // Resources.applycodechanges_6548_321 doesn't look good here
            _quickFixButton.Enabled = false;

            _navigatePreviousButton = (CommandBarButton)toolbar.Controls.Add(MsoControlType.msoControlButton, Temporary:true);
            _navigatePreviousButton.BeginGroup = true;
            _navigatePreviousButton.Caption = "Previous";
            _navigatePreviousButton.TooltipText = "Navigate to previous issue";
            _navigatePreviousButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _navigatePreviousButton.FaceId = 41; // Resources.112_LeftArrowLong_Blue_16x16_72 makes a gray block when disabled
            _navigatePreviousButton.Enabled = false;

            _navigateNextButton = (CommandBarButton)toolbar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            _navigateNextButton.Caption = "Next";
            _navigateNextButton.TooltipText = "Navigate to next issue";
            _navigateNextButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _navigateNextButton.FaceId = 39; // Resources.112_RightArrowLong_Blue_16x16_72 makes a gray block when disabled
            _navigateNextButton.Enabled = false;

            _refreshButton.Click += _refreshButton_Click;
            _quickFixButton.Click += _quickFixButton_Click;
            _navigatePreviousButton.Click += _navigatePreviousButton_Click;
            _navigateNextButton.Click += _navigateNextButton_Click;
        }

        private void _navigateNextButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_issues.Length == 0)
            {
                return;
            }

            if (_currentIssue == _issues.Length - 1)
            {
                _currentIssue = - 1;
            }

            _currentIssue++;
            OnNavigateCodeIssue(null, new NavigateCodeEventArgs(_issues[_currentIssue].QualifiedSelection.QualifiedName, _issues[_currentIssue].Context));
        }

        private void _navigatePreviousButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_issues.Length == 0)
            {
                return;
            }

            if (_currentIssue == 0)
            {
                _currentIssue = _issues.Length;
            }

            _currentIssue--;
            OnNavigateCodeIssue(null, new NavigateCodeEventArgs(_issues[_currentIssue].QualifiedSelection.QualifiedName, _issues[_currentIssue].Context));
        }

        private void OnNavigateCodeIssue(object sender, NavigateCodeEventArgs e)
        {
            try
            {
                var location = _vbe.FindInstruction(e.QualifiedName, e.Selection);
                location.CodeModule.CodePane.SetSelection(location.Selection);

                var codePane = location.CodeModule.CodePane;
                var selection = location.Selection;
                codePane.SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
                codePane.ForceFocus();
                SetQuickFixTooltip();
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.Assert(false);
            }
        }

        private void _quickFixButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                var fix = _issues[_currentIssue].GetQuickFixes().FirstOrDefault();
                if (!string.IsNullOrEmpty(fix.Key))
                {
                    fix.Value(_vbe);
                    _refreshButton_Click(null, ref CancelDefault);
                    _navigateNextButton_Click(null, ref CancelDefault);
                }
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.Assert(false);
            }
        }

        private void _refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            RefreshAsync();
        }

        private async void RefreshAsync()
        {
            var code = (_parser.Parse(_vbe.ActiveVBProject)).ToList();

            var results = new ConcurrentBag<CodeInspectionResultBase>();
            var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow);
            Parallel.ForEach(inspections, inspection =>
            {
                var result = inspection.GetInspectionResults(code);
                foreach (var inspectionResult in result)
                {
                    results.Add(inspectionResult);
                }
            });

            _issues = results.ToArray();
            _currentIssue = 0;

            var hasIssues = results.Any();
            _quickFixButton.Enabled = hasIssues;
            SetQuickFixTooltip();
            _navigateNextButton.Enabled = hasIssues;
            _navigatePreviousButton.Enabled = hasIssues;
            _statusButton.Caption = string.Format("{0} issue" + (results.Count == 1 ? string.Empty : "s"), results.Count);
        }

        private void SetQuickFixTooltip()
        {
            if (_issues.Length == 0)
            {
                _quickFixButton.TooltipText = string.Empty;
                return;
            }

            var fix = _issues[_currentIssue].GetQuickFixes().FirstOrDefault();
            if (string.IsNullOrEmpty(fix.Key))
            {
                _quickFixButton.Enabled = false;
            }

            _quickFixButton.TooltipText = fix.Key;
        }
    }
}