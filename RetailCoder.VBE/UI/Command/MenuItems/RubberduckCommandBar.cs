using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using ParserState = Rubberduck.Parsing.VBA.ParserState;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RubberduckCommandBar
    {
        private readonly RubberduckParserState _state;
        private readonly VBE _vbe;
        private readonly IShowParserErrorsCommand _command;

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;
        private CommandBarButton _selectionButton;

        public RubberduckCommandBar(RubberduckParserState state, VBE vbe, IShowParserErrorsCommand command)
        {
            _state = state;
            _vbe = vbe;
            _command = command;
            _state.StateChanged += State_StateChanged;
            Initialize();
        }

        private void _statusButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_state.Status == ParserState.Error)
            {
                _command.Execute(null);
            }
        }

        public void SetStatusText(string value = null)
        {
            Debug.WriteLine(string.Format("RubberduckCommandBar status text changes to '{0}'.", value));
            UiDispatcher.Invoke(() => _statusButton.Caption = value ?? RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status));
        }

        public void SetSelectionText(Declaration declaration)
        {
            if (declaration == null && _vbe.ActiveCodePane != null)
            {
                var selection = _vbe.ActiveCodePane.GetSelection();
                SetSelectionText(selection);
                _selectionButton.TooltipText = _selectionButton.Caption;
            }
            else if (declaration != null && !declaration.IsBuiltIn && declaration.DeclarationType != DeclarationType.Class && declaration.DeclarationType != DeclarationType.Module)
            {
                _selectionButton.Caption = string.Format("{0} ({1}): {2} ({3})", 
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.QualifiedSelection.Selection,
                    declaration.IdentifierName,
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType));
                _selectionButton.TooltipText = string.IsNullOrEmpty(declaration.DescriptionString)
                    ? _selectionButton.Caption
                    : declaration.DescriptionString;
            }
            else if (declaration != null)
            {
                var selection = _vbe.ActiveCodePane.GetSelection();
                _selectionButton.Caption = string.Format("{0}: {1} ({2}) {3}",
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.IdentifierName,
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType),
                    selection.Selection);
                _selectionButton.TooltipText = string.IsNullOrEmpty(declaration.DescriptionString)
                    ? _selectionButton.Caption
                    : declaration.DescriptionString;
            }
        }

        private void SetSelectionText(QualifiedSelection selection)
        {
            UiDispatcher.Invoke(() => _selectionButton.Caption = selection.ToString());
        }

        private void State_StateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("RubberduckCommandBar handles StateChanged...");
            SetStatusText(RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status));
        }

        public event EventHandler Refresh;

        private void OnRefresh()
        {
            var handler = Refresh;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        public void Initialize()
        {
            var commandbar = _vbe.CommandBars.Add("Rubberduck", MsoBarPosition.msoBarTop, false, true);

            _refreshButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            ParentMenuItemBase.SetButtonImage(_refreshButton, Resources.arrow_circle_double, Resources.arrow_circle_double_mask);
            _refreshButton.Style = MsoButtonStyle.msoButtonIcon;
            _refreshButton.Tag = "Refresh";
            _refreshButton.TooltipText =RubberduckUI.RubberduckCommandbarRefreshButtonTooltip;
            _refreshButton.Click += refreshButton_Click;

            _statusButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            _statusButton.Style = MsoButtonStyle.msoButtonCaption;
            _statusButton.Tag = "Status";
            _statusButton.Click += _statusButton_Click;

            _selectionButton = (CommandBarButton)commandbar.Controls.Add(MsoControlType.msoControlButton);
            _selectionButton.Style = MsoButtonStyle.msoButtonCaption;
            _selectionButton.BeginGroup = true;
            _selectionButton.Enabled = false;

            commandbar.Visible = true;
        }

        private void refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRefresh();
        }
    }
}
