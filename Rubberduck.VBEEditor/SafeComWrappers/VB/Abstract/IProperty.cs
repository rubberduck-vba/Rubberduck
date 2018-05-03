using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IProperty : ISafeComWrapper, IEquatable<IProperty>
    {
        string Name { get; }
        object Value { get; set; }
        object Object { get; set; }

        int IndexCount { get; }
        object GetIndexedValue(object index1, object index2 = null, object index3 = null, object index4 = null);
        void SetIndexedValue(object value, object index1, object index2 = null, object index3 = null, object index4 = null);

        IProperties Collection { get; }
        IProperties Parent { get; }
        IApplication Application { get; }
        IVBE VBE { get; }
    }
}