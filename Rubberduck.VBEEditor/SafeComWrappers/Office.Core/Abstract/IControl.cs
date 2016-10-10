using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface IControl : ISafeComWrapper, IEquatable<IControl>
    {
        string Name { get; set; }
    }
}