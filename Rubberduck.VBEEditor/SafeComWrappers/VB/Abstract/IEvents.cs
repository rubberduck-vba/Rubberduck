using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    // This interface is not included on IVBE, as it is only safe in VB6
    // https://stackoverflow.com/questions/41055765/whats-the-difference-between-commandbarevents-click-and-commandbarbutton-click/41066408#41066408
    public interface IEvents : ISafeComWrapper, IEquatable<IEvents>
    {
        ICommandBarEvents CommandBarEvents { get; }
    }
}
