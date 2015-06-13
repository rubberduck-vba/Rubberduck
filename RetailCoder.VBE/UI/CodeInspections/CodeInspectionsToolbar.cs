using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.VBIDEApi;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Properties;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsToolbar : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IEnumerable<IInspection> _inspections;
        private readonly IRubberduckParser _parser;
        private readonly IInspector _inspector;

        private IList<ICodeInspectionResult> _issues;
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
            _toolbar = _vbe.CommandBars.Add(RubberduckUI.CodeInspections, null, null, true);
            _refreshButton = (CommandBarButton)_toolbar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            _refreshButton.TooltipText = RubberduckUI.CodeInspections_Run;

            var refreshIcon = Resources.Refresh;
            refreshIcon.MakeTransparent(Color.Magenta);
            Menu.SetButtonImage(_refreshButton, refreshIcon);

            _statusButton = (CommandBarButton)_toolbar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            _statusButton.Caption = string.Format(RubberduckUI.CodeInspections_NumberOfIssues, 0, "s");
            _statusButton.FaceId = 463; // Resources.Warning doesn't look good here
            _statusButton.Style = MsoButtonStyle.msoButtonIconAndCaption;

            _quickFixButton = (CommandBarButton)_toolbar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            _quickFixButton.Caption = RubberduckUI.Fix;
            _quickFixButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _quickFixButton.FaceId = 305; // Resources.applycodechanges_6548_321 doesn't look good here
            _quickFixButton.Enabled = false;

            _navigatePreviousButton = (CommandBarButton)_toolbar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            _navigatePreviousButton.BeginGroup = true;
            _navigatePreviousButton.Caption = RubberduckUI.Previous;
            _navigatePreviousButton.TooltipText = RubberduckUI.Previous;
            _navigatePreviousButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _navigatePreviousButton.FaceId = 41; // Resources.112_LeftArrowLong_Blue_16x16_72 makes a gray Block when disabled
            _navigatePreviousButton.Enabled = false;

            _navigateNextButton = (CommandBarButton)_toolbar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            _navigateNextButton.Caption = RubberduckUI.Next;
            _navigateNextButton.TooltipText = RubberduckUI.Next;
            _navigateNextButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            _navigateNextButton.FaceId = 39; // Resources.112_RightArrowLong_Blue_16x16_72 makes a gray Block when disabled
            _navigateNextButton.Enabled = false;

            _refreshButton.ClickEvent += _refreshButton_Click;
            _quickFixButton.ClickEvent += _quickFixButton_Click;
            _navigatePreviousButton.ClickEvent += _navigatePreviousButton_Click;
            _navigateNextButton.ClickEvent += _navigateNextButton_Click;

            _inspector.IssuesFound += OnIssuesFound;
            _inspector.Reset += OnReset;
            _inspector.ParseCompleted += _inspector_ParseCompleted;
        }

        private IEnumerable<VBProjectParseResult> _parseResults;

        private void _inspector_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            if (sender != this)
            {
                return;
            }

            _parseResults = e.ParseResults;
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
                    fix.Value();
                    _refreshButton_Click(null, ref CancelDefault);
                    _navigateNextButton_Click(null, ref CancelDefault);
                }
            }
            catch (Exception exception)
            {
                Debug.Assert(false);
            }
        }

        private CancellationTokenSource _tokenSource;
        private CommandBar _toolbar;

        public bool ToolbarVisible
        {
            get { return _toolbar.Visible; }
            set { _toolbar.Visible = value; }
        }

        public Point ToolbarCoords
        {
            get { return new Point(_toolbar.Top, _toolbar.Left); }
            set
            {
                _toolbar.Top = value.X;
                _toolbar.Left = value.Y;
            }
        }

        private void _refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _tokenSource = new CancellationTokenSource();
            var token = _tokenSource.Token;
            RefreshAsync(token);
        }

        private void OnIssuesFound(object sender, InspectorIssuesFoundEventArg e)
        {
            _issueCount = _issueCount + e.Issues.Count;
            _statusButton.Caption = string.Format(RubberduckUI.CodeInspections_NumberOfIssues, _issueCount, (_issueCount == 1 ? string.Empty : "s"));
        }

        private async void RefreshAsync(CancellationToken token)
        {
            try
            {
                var projectParseResult = await _inspector.Parse(_vbe.ActiveVBProject, this);
                _issues = await _inspector.FindIssuesAsync(projectParseResult, token);
            }
            catch (COMException)
            {
                // burp
            }

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

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            _refreshButton.ClickEvent -= _refreshButton_Click;
            _quickFixButton.ClickEvent -= _quickFixButton_Click;
            _navigatePreviousButton.ClickEvent -= _navigatePreviousButton_Click;
            _navigateNextButton.ClickEvent -= _navigateNextButton_Click;  

            _toolbar.Delete();
        }
    }
}