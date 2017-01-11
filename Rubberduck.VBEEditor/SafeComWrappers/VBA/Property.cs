using System.Diagnostics.CodeAnalysis;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    [SuppressMessage("ReSharper", "UseIndexedProperty")]
    public class Property : SafeComWrapper<VB.Property>, IProperty
    {
        public Property(VB.Property target) 
            : base(target)
        {
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
        }

        public int IndexCount
        {
            get { return IsWrappingNullReference ? 0 : Target.NumIndices; }
        }

        public IProperties Collection
        {
            get { return new Properties(IsWrappingNullReference ? null : Target.Collection); }
        }

        public IProperties Parent
        {
            get { return new Properties(IsWrappingNullReference ? null : Target.Parent); }
        }

        public IApplication Application
        {
            get { return new Application(IsWrappingNullReference ? null : Target.Application); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public object Value
        {
            get { return IsWrappingNullReference ? null : Target.Value; }
            set { if (!IsWrappingNullReference) Target.Value = value; }
        }

        /// <summary>
        /// Getter can return an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object GetIndexedValue(object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            return IsWrappingNullReference ? null : Target.get_IndexedValue(index1, index2, index3, index4);
        }

        public void SetIndexedValue(object value, object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            if (!IsWrappingNullReference) Target.set_IndexedValue(index1, index2, index3, index4, value);
        }

        /// <summary>
        /// Getter returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Object
        {
            get { return IsWrappingNullReference ? null : Target.Object; }
            set { if (!IsWrappingNullReference) Target.Object = value; }
        }

        public override bool Equals(ISafeComWrapper<VB.Property> other)
        {
            return IsEqualIfNull(other) ||
                (other != null && other.Target.Name == Name && ReferenceEquals(other.Target.Parent, Target.Parent));
        }

        public bool Equals(IProperty other)
        {
            return Equals(other as SafeComWrapper<VB.Property>);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(Name, IndexCount, Parent.Target);
        }
    }
}