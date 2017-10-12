using System.Diagnostics.CodeAnalysis;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VB6
{
    [SuppressMessage("ReSharper", "UseIndexedProperty")]
    public class Property : SafeComWrapper<VB6IA.Property>, IProperty
    {
        public Property(VB6IA.Property target) 
            : base(target)
        {
        }

        public string Name => IsWrappingNullReference ? string.Empty : Target.Name;

        public int IndexCount => IsWrappingNullReference ? 0 : Target.NumIndices;

        public IProperties Collection => new Properties(IsWrappingNullReference ? null : Target.Collection);

        public IProperties Parent => new Properties(IsWrappingNullReference ? null : Target.Parent);

        public IApplication Application => new Application((VB6IA.Application) (IsWrappingNullReference ? null : Target.Application));

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public object Value
        {
            get => IsWrappingNullReference ? null : Target.Value;
            set => Target.Value = value;
        }

        /// <summary>
        /// Getter can return an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object GetIndexedValue(object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            return Target.get_IndexedValue(index1, index2, index3, index4);
        }

        public void SetIndexedValue(object value, object index1, object index2 = null, object index3 = null, object index4 = null)
        {
            Target.set_IndexedValue(index1, index2, index3, index4, value);
        }

        /// <summary>
        /// Getter returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Object
        {
            get => IsWrappingNullReference ? null : Target.Object;
            set => Target.Object = value;
        }

        public override bool Equals(ISafeComWrapper<VB6IA.Property> other)
        {
            return IsEqualIfNull(other) ||
                (other != null && other.Target.Name == Name && ReferenceEquals(other.Target.Parent, Target.Parent));
        }

        public bool Equals(IProperty other)
        {
            return Equals(other as SafeComWrapper<VB6IA.Property>);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(Name, IndexCount, Parent.Target);
        }
    }
}