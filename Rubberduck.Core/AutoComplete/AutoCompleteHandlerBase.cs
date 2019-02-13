using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete
{
    /// <summary>
    /// A base class/interface for AC services / "handlers".
    /// </summary>
    public abstract class AutoCompleteHandlerBase
    {
        protected AutoCompleteHandlerBase(ICodePaneHandler pane)
        {
            CodePaneHandler = pane;
        }

        protected ICodePaneHandler CodePaneHandler { get; }

        /// <summary>
        /// A method that returns <c>false</c> if the input isn't handled, <c>true</c> if it is.
        /// </summary>
        /// <param name="e">The autocompletion event info</param>
        /// <param name="settings">The current AC settings</param>
        /// <param name="result">If handled, the resulting <c>CodeString</c></param>
        /// <returns></returns>
        public abstract bool Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings, out CodeString result);
    }
}