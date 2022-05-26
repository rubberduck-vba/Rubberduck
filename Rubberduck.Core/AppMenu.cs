using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
            InitializeRubberduckCommandBar();
            InitializeRubberduckMenus();
        }

        private void InitializeRubberduckCommandBar()
        {
            try
            {
                _stateBar.Initialize();
            }
            catch (COMException exception)
            {
                _logger.Error(exception);
                throw; // NOTE: this exception should bubble up to _Extension.Startup() and cleanly fail the add-in's initialization.
            }
            catch (Exception exception)
            {
                // we don't want to abort init just because some CanExecute method threw a NRE
                _logger.Error(exception);
            }
        }

        private void InitializeRubberduckMenus()
        { 
            foreach (var menu in _menus)
            {
                try
                {
                    menu.Initialize();
                }
                catch (COMException exception)
                {
                    _logger.Error(exception);
                    throw; // NOTE: this exception should bubble up to _Extension.Startup() and cleanly fail the add-in's initialization.
                }
                catch (Exception exception)
                {
                    // we don't want to abort init just because some CanExecute method threw a NRE
                    _logger.Error(exception);
                }
            }
            EvaluateCanExecuteAsync(_parser.State, CancellationToken.None).Wait();
        }

        public async void OnSelectedDeclarationChange(object sender, DeclarationChangedEventArgs e)
        {
            await EvaluateCanExecuteAsync(_parser.State, CancellationToken.None);
        }

        private async void OnParserStateChanged(object sender, EventArgs e)
        {            
            await EvaluateCanExecuteAsync(_parser.State, CancellationToken.None);
        }

        public async Task EvaluateCanExecuteAsync(RubberduckParserState state, CancellationToken token)
        {
            foreach (var menu in _menus)
            {
                try
                {
                    await menu.EvaluateCanExecuteAsync(state, token);
                }
                catch (Exception exception)
                {
                    // swallow exception to evaluate the other commands
                    _logger.Error(exception);
                }
            }
        }

        public void Localize()
        {
            LocalizeRubberduckCommandBar();
            LocalizeRubberduckMenus();
        }

        private void LocalizeRubberduckCommandBar()
        {
            try
            {
                _stateBar.Localize();
                _stateBar.SetStatusLabelCaption(_parser.State.Status);
            }
            catch (Exception exception)
            {
                _logger.Error(exception);
            }
        }

        private void LocalizeRubberduckMenus()
        {
            foreach (var menu in _menus)
            {
                try
                {
                    menu.Localize();
                }
                catch (Exception exception)
                {
                    _logger.Error(exception);
                }
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
                try
                {
                    _logger.Debug($"Starting removal of top-level menu {menu.GetType()}.");
                    menu.RemoveMenu();
                    //We do this here and not in the menu items because we only want to dispose of/release the parents of the top level parent menus.
                    //The parents further down get disposed of/released as part of the remove chain.
                    _logger.Trace($"Removing parent menu of top-level menu {menu.GetType()}.");
                    menu.Parent.Dispose();
                    menu.Parent = null;
                }
                catch (Exception exception)
                {
                    _logger.Error(exception);
                }
            }
        }
    }
}
