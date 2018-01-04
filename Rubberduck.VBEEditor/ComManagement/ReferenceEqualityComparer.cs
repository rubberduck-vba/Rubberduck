using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace Rubberduck.VBEditor.ComManagement
{
    //See https://stackoverflow.com/a/41169463/5536802
    public sealed class ReferenceEqualityComparer : IEqualityComparer, IEqualityComparer<object>
    {
        public static ReferenceEqualityComparer Default { get; } = new ReferenceEqualityComparer();

        public new bool Equals(object x, object y) => ReferenceEquals(x, y);
        public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
