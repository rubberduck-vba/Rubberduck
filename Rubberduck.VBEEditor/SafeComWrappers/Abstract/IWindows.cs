using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IWindows : ISafeComWrapper, IComCollection<IWindow>, IEquatable<IWindows>
    {
        IVBE VBE { get; }
        IApplication Parent { get; }
        IWindow CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition, ref object docObj);
    }
}