using Microsoft.Office.Core;

namespace Rubberduck.UI.Commands
{
    public class RubberduckParentMenu : ParentMenu
    {
        private readonly CodeExplorerCommandMenuItem _codeExplorer;
        private readonly OptionsCommandMenuItem _options;
        private readonly AboutCommandMenuItem _about;

        public RubberduckParentMenu(CommandBarControls parent, int beforeIndex,
            CodeExplorerCommandMenuItem codeExplorer,
            OptionsCommandMenuItem options,
            AboutCommandMenuItem about)
            : base(parent, "RubberduckMenu", () => RubberduckUI.RubberduckMenu, beforeIndex)
        {
            _codeExplorer = codeExplorer;
            _options = options;
            _about = about;
        }

        public override void Initialize()
        {
            AddItem(_codeExplorer);
            AddItem(_options, true);
            AddItem(_about, true);
        }
    }
}