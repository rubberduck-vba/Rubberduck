using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteHandlerBase
    {
        protected AutoCompleteHandlerBase(ICodePaneHandler pane)
        {
            CodePaneHandler = pane;
        }

        protected ICodePaneHandler CodePaneHandler { get; }

        public abstract bool Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings, out CodeString result);
    }
}