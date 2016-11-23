using System;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class RubberduckCommandBar : AppCommandBarBase, IDisposable
    {
        public RubberduckCommandBar(IEnumerable<ICommandMenuItem> items) 
            : base("Rubberduck", CommandBarPosition.Top, items)
        {
        }

        public void SetStatusLabelCaption(ParserState state)
        {
            var caption = RubberduckUI.ResourceManager.GetString("ParserState_" + state, Settings.Settings.Culture);
            SetStatusLabelCaption(caption);
        }

        public void SetStatusLabelCaption(string caption)
        {
            var child = FindChildByTag(typeof(ShowParserErrorsCommandMenuItem).FullName) as ShowParserErrorsCommandMenuItem;
            if (child == null) { return; }

            UiDispatcher.Invoke(() => child.SetCaption(caption));
            Localize();
        }

        public void SetContextSelectionCaption(string caption)
        {
            var child = FindChildByTag(typeof(ContextSelectionLabelMenuItem).FullName) as ContextSelectionLabelMenuItem;
            if (child == null) { return; }

            UiDispatcher.Invoke(() => child.SetCaption(caption));
            Localize();
        }

        public void Dispose()
        {
            RemoveChildren();
            Item.Delete();
            Item.Release(true);
        }
    }

    public enum RubberduckCommandBarItemDisplayOrder
    {
        RequestReparse,
        ShowErrors,
        ContextStatus
    }
}