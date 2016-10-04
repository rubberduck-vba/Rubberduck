namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Property : SafeComWrapper<Microsoft.Vbe.Interop.Property>
    {
        public Property(Microsoft.Vbe.Interop.Property comObject) 
            : base(comObject)
        {
        }

        public object Value
        {
            get { return InvokeResult(() => ComObject.Value); }
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

        public int IndexCount { get { return InvokeResult(() => ComObject.NumIndices); } }
        public Application Application { get { return new Application(InvokeResult(() => ComObject.Application)); } }
        public Properties Parent { get { return new Properties(InvokeResult(() => ComObject.Parent)); } }
        public string Name { get { return InvokeResult(() => ComObject.Name); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public Properties Collection { get { return new Properties(InvokeResult(() => ComObject.Collection)); } }

        /// <summary>
        /// Getter returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Object
        {
            get { return InvokeResult(() => ComObject.Object); }
            set { Invoke(() => ComObject.Object = value); }
        }
    }
}