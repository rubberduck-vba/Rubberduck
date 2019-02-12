using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    /// <summary>
    /// Provides a generic manner of reporting document state for use within Rubberduck
    /// </summary>
    /// <remarks>
    /// Different hosts may have different states and they may behave differently. For example,
    /// Excel's Worksheet document has no true design-time state; only time it is locked is when
    /// the focus is inside the formula box in the Excel UI. On the other hand, Access has both
    /// a design time and a run time state for its Form document and Report document. The enum
    /// is meant to provide a generic representation and must correspond exactly to how the host
    /// will treat the document. All hosts need not implement the full set of the enum but must
    /// represent it exactly as per the description of the enum member.
    /// </remarks>
    public enum DocumentState
    {
        /// <summary>
        /// The document is not accessible. It might be closed or otherwise unavailable. In this case
        /// it should be presumed it is not safe to use the <see cref="IVBComponent"/> of that document.
        /// </summary>
        Inaccessible,
        /// <summary>
        /// The document is open in design mode.
        /// </summary>
        DesignView,
        /// <summary>
        /// The document is open in non-design mode. Not all design-time operations are available.
        /// </summary>
        ActiveView
    }
    
    public interface IHostDocument
    {
        QualifiedModuleName QualifiedName { get; }
        string DocumentName { get; }
        string ClassName { get; }
        DocumentState State { get; }
        bool TryGetTarget(out SafeIDispatchWrapper iDispatch);
    } 

    public class HostDocument : IHostDocument
    {
        private readonly Func<SafeIDispatchWrapper> _getTargetFunc;

        public HostDocument(QualifiedModuleName qualifedName, string name, string className, DocumentState state, Func<SafeIDispatchWrapper> getTargetFunc)
        {
            QualifiedName = qualifedName;
            DocumentName = name;
            ClassName = className;
            State = state;

            _getTargetFunc = getTargetFunc;
        }

        public QualifiedModuleName QualifiedName { get; }
        public string DocumentName { get; }
        public string ClassName { get; }
        public DocumentState State { get; }

        public bool TryGetTarget(out SafeIDispatchWrapper iDispatch)
        {
            if (_getTargetFunc == null)
            {
                iDispatch = null;
                return false;
            }

            try
            {
                iDispatch = _getTargetFunc.Invoke();
                return true;
            }
            catch
            {
                iDispatch = null;
                return false;
            }
        }
    }
}