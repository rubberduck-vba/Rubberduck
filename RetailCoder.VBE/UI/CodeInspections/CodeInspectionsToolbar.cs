using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Properties;
using Rubberduck.VBA;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsToolbar
    {
        private readonly VBE _vbe;
        private readonly IEnumerable<IInspection> _inspections;
        private readonly IRubberduckParser _parser;
        private readonly IInspector _inspector;

        private List<ICodeInspectionResult> _issues;
        private int _currentIssue;
        private int _issueCount;

        public CodeInspectionsToolbar(VBE vbe, IRubberduckParser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _parser = parser;
            _inspections = inspections;
        }

        public CodeInspectionsToolbar(VBE vbe, IInspector inspector)
        {
            _vbe = vbe;
            _inspector = inspector;
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
            _navigatePreviousButton.FaceId = 41; // Resources.112_LeftArrowLong_Blue_16x16_72 makes a gray Block when disabled
            _navigatePreviousButton.Enabled = false;

            _navigateNextButton = (CommandBarButton)toolbar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
            _navigateNextButton.Caption = "Next";
            _navigateNextButton.TooltipText = "Navigate to next issue";
            _navigateNextButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _navigateNextButton.FaceId = 39; // Resources.112_RightArrowLong_Blue_16x16_72 makes a gray Block when disabled
            _navigateNextButton.Enabled = false;

            _refreshButton.Click += _refreshButton_Click;
            _quickFixButton.Click += _quickFixButton_Click;
            _navigatePreviousButton.Click += _navigatePreviousButton_Click;
            _navigateNextButton.Click += _navigateNextButton_Click;

            _inspector.IssuesFound += OnIssuesFound;
            _inspector.Reset += OnReset;
            _inspector.ParseCompleted += _inspector_ParseCompleted;
        }

        private IEnumerable<VBProjectParseResult> _parseResults;
        
        // indicates that the _parseResults are no longer in sync with UI
        private bool _needsResync;

        private void _inspector_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            if (sender == this)
            {
                _needsResync = false;
                _parseResults = e.ParseResults;
                Task.Run(() => RefreshAsync());
            }
            else
            {
                _parseResults = e.ParseResults;
                _needsResync = true;
            }
        }

        private void _navigateNextButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_issues.Count == 0)
            {
                return;
            }

            if (_currentIssue == _issues.Count - 1)
            {
                _currentIssue = - 1;
            }

            _currentIssue++;
            OnNavigateCodeIssue(null, new NavigateCodeEventArgs(_issues[_currentIssue].QualifiedSelection.QualifiedName, _issues[_currentIssue].Context));
        }

        private void _navigatePreviousButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_issues.Count == 0)
            {
                return;
            }

            if (_currentIssue == 0)
            {
                _currentIssue = _issues.Count;
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
                Debug.Assert(false);
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
                Debug.Assert(false);
            }
        }

        private void _refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            RefreshAsync();
        }

        private void OnIssuesFound(object sender, InspectorIssuesFoundEventArg e)
        {
            _issueCount = _issueCount + e.Issues.Count;
            _statusButton.Caption = string.Format("{0} issue" + (_issueCount == 1 ? string.Empty : "s"), _issueCount);
        }

        private async void RefreshAsync()
        {
            if (_parseResults == null || !_needsResync)
            {
                _inspector.Parse(_vbe, this);
                return;
            }

            var parseResults = _parseResults.SingleOrDefault(p => p.Project == _vbe.ActiveVBProject);
            if (parseResults == null)
            {
                return;
            }

            var result = await _inspector.FindIssuesAsync(parseResults);
            _issues = result.ToList();

            var hasIssues = _issues.Any();
            _quickFixButton.Enabled = hasIssues;
            SetQuickFixTooltip();
            _navigateNextButton.Enabled = hasIssues;
            _navigatePreviousButton.Enabled = hasIssues;
        }

        private void OnReset(object sender, EventArgs e)
        {
            _currentIssue = 0;
            _issueCount = 0;
        }

        private void SetQuickFixTooltip()
        {
            if (_issues.Count == 0)
            {
                _quickFixButton.TooltipText = string.Empty;
                _statusButton.TooltipText = string.Empty;
                return;
            }

            var fix = _issues[_currentIssue].GetQuickFixes().FirstOrDefault();
            if (string.IsNullOrEmpty(fix.Key))
            {
                _quickFixButton.Enabled = false;
            }

            _quickFixButton.TooltipText = fix.Key;
            _statusButton.TooltipText = _issues[_currentIssue].Name;
        }
    }
}