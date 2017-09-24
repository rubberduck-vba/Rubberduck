using System;
using System.Collections;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IDesigner : ISafeComWrapper, IEquatable<IVBComponent>
    { 
        IList Selected { get; }
    }
}
