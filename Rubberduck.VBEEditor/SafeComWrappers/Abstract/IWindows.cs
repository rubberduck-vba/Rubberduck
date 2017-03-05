using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public struct ToolWindowInfo
    {
        private readonly IWindow _window;
        private readonly object _control;

        public ToolWindowInfo(IWindow window, object control) : this()
        {
            _window = window;
            _control = control;
        }

        public IWindow ToolWindow { get { return _window; } }
        public object UserControl { get { return _control; } }
    }

    public interface IWindows : ISafeComWrapper, IComCollection<IWindow>, IEquatable<IWindows>
    {
        IVBE VBE { get; }
        IApplication Parent { get; }
        ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition);
    }
}