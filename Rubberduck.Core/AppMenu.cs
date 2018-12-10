using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.CommandBars;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck
{
    public class AppMenu : IAppMenu, IDisposable
    {
        private readonly IReadOnlyList<IParentMenuItem> _menus;
        private readonly IParseCoordinator _parser;
        private readonly ISelectionChangeService _selectionService;
        private readonly RubberduckCommandBar _stateBar;

        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public AppMenu(IEnumerable<IParentMenuItem> menus, IParseCoordinator parser, ISelectionChangeService selectionService, RubberduckCommandBar stateBar)
        {
            _menus = menus.ToList();
            _parser = parser;
            _selectionService = selectionService;
            _stateBar = stateBar;

            _parser.State.StateChanged += OnParserStateChanged;
            _selectionService.SelectedDeclarationChanged += OnSelectedDeclarationChange;
        }

        public void Initialize()
        {
            _stateBar.Initialize();
            foreach (var menu in _menus)
            {
                menu.Initialize();
            }
            EvaluateCanExecute(_parser.State);
        }

        public void OnSelectedDeclarationChange(object sender, DeclarationChangedEventArgs e)
        {
            EvaluateCanExecute(_parser.State);
        }

        private void OnParserStateChanged(object sender, EventArgs e)
        {            
            EvaluateCanExecute(_parser.State);
        }

        public void EvaluateCanExecute(RubberduckParserState state)
        {
            foreach (var menu in _menus)
            {
                menu.EvaluateCanExecute(state);
            }
        }

        public void Localize()
        {
            _stateBar.Localize();
            _stateBar.SetStatusLabelCaption(_parser.State.Status);

            foreach (var menu in _menus)
            {
                menu.Localize();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }
            _isDisposed = true;

            _parser.State.StateChanged -= OnParserStateChanged;
            _selectionService.SelectedDeclarationChanged -= OnSelectedDeclarationChange;

            RemoveMenus();
        }

        private void RemoveMenus()
        {
            foreach (var menu in _menus.Where(menu => menu.Item != null))
            {
                _logger.Debug($"Starting removal of top-level menu {menu.GetType()}.");
                menu.RemoveMenu();
                //We do this here and not in the menu items because we only want to dispose of/release the parents of the top level parent menus.
                //The parents further down get disposed of/released as part of the remove chain.
                _logger.Trace($"Removing parent menu of top-level menu {menu.GetType()}.");
                menu.Parent.Dispose();
                menu.Parent = null;
            }
        }
    }
}
