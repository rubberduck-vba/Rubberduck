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

        public void SetStatusLabelCaption(ParserState state, int? errorCount = null)
        {
            var caption = RubberduckUI.ResourceManager.GetString("ParserState_" + state, Settings.Settings.Culture);
            SetStatusLabelCaption(caption, errorCount);
        }

        public void SetStatusLabelCaption(string caption, int? errorCount = null)
        {
            var reparseCommandButton = FindChildByTag(typeof(ReparseCommandMenuItem).FullName) as ReparseCommandMenuItem;
            if (reparseCommandButton == null) { return; }

            var showErrorsCommandButton = FindChildByTag(typeof(ShowParserErrorsCommandMenuItem).FullName) as ShowParserErrorsCommandMenuItem;
            if (showErrorsCommandButton == null) { return; }

            UiDispatcher.Invoke(() =>
            {
                reparseCommandButton.SetCaption(caption);
                reparseCommandButton.SetToolTip(string.Format(RubberduckUI.ReparseToolTipText, caption));
                if (errorCount.HasValue && errorCount.Value > 0)
                {
                    showErrorsCommandButton.SetToolTip(string.Format(RubberduckUI.ParserErrorToolTipText, errorCount.Value));
                }
            });
            Localize();
        }

        public void SetContextSelectionCaption(string caption)
        {
            var child = FindChildByTag(typeof(ContextSelectionLabelMenuItem).FullName) as ContextSelectionLabelMenuItem;
            if (child == null) { return; }

            UiDispatcher.Invoke(() =>
            {
                child.SetCaption(caption);
                //child.SetToolTip(?);
            });
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