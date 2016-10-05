using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    [SuppressMessage("ReSharper", "UseIndexedProperty")]
    public class Property : SafeComWrapper<Microsoft.Vbe.Interop.Property>, IEquatable<Property>
    {
        public Property(Microsoft.Vbe.Interop.Property comObject) 
            : base(comObject)
        {
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Name); }
        }

        public int IndexCount
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.NumIndices); }
        }

        public Properties Collection
        {
            get { return new Properties(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Collection)); }
        }

        public Properties Parent
        {
            get { return new Properties(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Parent)); }
        }

        public Application Application
        {
            get { return new Application(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Application)); }
        }

        public VBE VBE
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.VBE)); }
        }

        public object Value
        {
            get { return IsWrappingNullReference ? null : InvokeResult(() => ComObject.Value); }
            set { Invoke(() => ComObject.Value = value); }
        }

        /// <summary>
        /// Getter can return an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object GetIndexedValue(object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            return InvokeResult(() => ComObject.get_IndexedValue(index1, index2, index3, index4));
        }

        public void SetIndexedValue(object value, object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            Invoke(() => ComObject.set_IndexedValue(index1, index2, index3, index4, value));
        }

        /// <summary>
        /// Getter returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Object
        {
            get { return IsWrappingNullReference ? null : InvokeResult(() => ComObject.Object); }
            set { Invoke(() => ComObject.Object = value); }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(ComObject);
            } 
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Property> other)
        {
            return IsEqualIfNull(other) ||
                (other != null && other.ComObject.Name == Name && ReferenceEquals(other.ComObject.Parent, ComObject.Parent));
        }

        public bool Equals(Property other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Property>);
        }

        public override int GetHashCode()
        {
            return ComputeHashCode(Name, IndexCount, Parent.ComObject);
        }
    }
}