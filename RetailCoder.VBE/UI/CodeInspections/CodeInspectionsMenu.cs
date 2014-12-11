using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.Properties;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class CodeInspectionsMenu
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IEnumerable<IInspection> _inspections;
        private readonly Parser _parser;

        public CodeInspectionsMenu(VBE vbe, AddIn addin, Parser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
            _inspections = inspections;
        }

        private CommandBarButton _codeInspectionsButton;

        public void Initialize(CommandBarControls menuControls)
        {
            _codeInspectionsButton = menuControls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            Debug.Assert(_codeInspectionsButton != null);

            _codeInspectionsButton.Caption = "Code &Inspections";

            _codeInspectionsButton.Click += OnCodeInspectionsButtonClick;
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            var presenter = new CodeInspectionsDockablePresenter(_parser, _inspections, _vbe, _addin);
            presenter.Show();
        }
    }

    [ComVisible(false)]
    public class CodeInspectionsToolbar
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IEnumerable<IInspection> _inspections;
        private readonly Parser _parser;

        public CodeInspectionsToolbar(VBE vbe, AddIn addin, Parser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
            _inspections = inspections;
        }

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;
        private CommandBarButton _quickFixButton;
        private CommandBarButton _navigatePreviousButton;
        private CommandBarButton _navigateNextButton;
        private CommandBarButton _showInspectionsWindowButton;

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
            _quickFixButton.TooltipText = "Address the issue";
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
        }
    }
}
