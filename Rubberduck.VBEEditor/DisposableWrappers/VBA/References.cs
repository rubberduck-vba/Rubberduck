using System;
using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class References : SafeComWrapper<Microsoft.Vbe.Interop.References>, IEnumerable<Reference>
    {
        public References(Microsoft.Vbe.Interop.References comObject) 
            : base(comObject)
        {
            comObject.ItemAdded += comObject_ItemAdded;
            comObject.ItemRemoved += comObject_ItemRemoved;
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved; 

        private void comObject_ItemRemoved(Microsoft.Vbe.Interop.Reference reference)
        {
            var handler = ItemRemoved;
            if (handler == null) { return; }
            handler.Invoke(this, new ReferenceEventArgs(new Reference(reference)));
        }

        private void comObject_ItemAdded(Microsoft.Vbe.Interop.Reference reference)
        {
            var handler = ItemAdded;
            if (handler == null) { return; }
            handler.Invoke(this, new ReferenceEventArgs(new Reference(reference)));
        }

        public Reference Item(object index)
        {
            return new Reference(InvokeResult(() => ComObject.Item(index)));
        }

        public Reference AddFromGuid(string guid, int major, int minor)
        {
            return new Reference(InvokeResult(() => ComObject.AddFromGuid(guid, major, minor)));
        }

        public Reference AddFromFile(string path)
        {
            return new Reference(InvokeResult(() => ComObject.AddFromFile(path)));
        }

        public void Remove(Reference reference)
        {
            Invoke(() => ComObject.Remove(reference.ComObject));
        }

        public VBProject Parent { get { return new VBProject(InvokeResult(() => ComObject.Parent)); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        IEnumerator<Reference> IEnumerable<Reference>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.References, Reference>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.References, Reference>(ComObject);
        }
    }
}