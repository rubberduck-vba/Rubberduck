using System;
using System.Collections.Generic;
using System.Linq;
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
            _parser.State.StateChanged -= OnParserStateChanged;
            _selectionService.SelectedDeclarationChanged -= OnSelectedDeclarationChange;

            // note: doing this wrecks the teardown process. counter-intuitive? sure. but hey it works.
            //foreach (var menu in _menus.Where(menu => menu.Item != null))
            //{
            //    menu.RemoveChildren();
            //    menu.Item.Delete();
            //}
        }
    }
}
