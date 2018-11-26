using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public struct ToolWindowInfo
    {
        public ToolWindowInfo(IWindow window, object control) : this()
        {
            ToolWindow = window;
            UserControl = control;
        }

        public IWindow ToolWindow { get; }

        public object UserControl { get; }
    }

    public interface IWindows : ISafeComWrapper, IComCollection<IWindow>, IEquatable<IWindows>
    {
        IVBE VBE { get; }
        IApplication Parent { get; }
        ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition);
        void ReleaseDockableHosts();
    }
}