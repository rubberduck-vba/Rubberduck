using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICodePane : ISafeComWrapper, IEquatable<ICodePane>
    {
        IVBE VBE { get; }
        ICodePanes Collection { get; }
        IWindow Window { get; }
        int TopLine { get; set; }
        int CountOfVisibleLines { get; }
        ICodeModule CodeModule { get; }
        CodePaneView CodePaneView { get; }
        Selection Selection { get; set; }
        QualifiedSelection? GetQualifiedSelection();
        void Show();
    }
}