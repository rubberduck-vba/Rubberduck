using System;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICodePane : ISubclassAttachable, IEquatable<ICodePane>
    {
        IVBE VBE { get; }
        ICodePanes Collection { get; }
        IWindow Window { get; }
        int TopLine { get; set; }
        int CountOfVisibleLines { get; }
        ICodeModule CodeModule { get; }
        CodePaneView CodePaneView { get; }
        /// <summary>
        /// Gets or sets a 1-based <see cref="Selection"/> representing the current selection in the code pane.
        /// </summary>
        Selection Selection { get; set; }
        QualifiedSelection? GetQualifiedSelection();
        QualifiedModuleName QualifiedModuleName { get; }
        void Show();
    }
}