using System.Collections;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class VBComponents : SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>, IEnumerable<VBComponent>
    {
        public VBComponents(Microsoft.Vbe.Interop.VBComponents comObject) 
            : base(comObject)
        {
        }

        public void Remove(VBComponent item)
        {
            Invoke(() => ComObject.Remove(item.ComObject));
        }

        public VBComponent Add(ComponentType type)
        {
            return new VBComponent(InvokeResult(() => ComObject.Add((vbext_ComponentType)type)));
        }

        public VBComponent Import(string path)
        {
            return new VBComponent(InvokeResult(() => ComObject.Import(path)));
        }

        public VBProject Parent { get { return new VBProject(InvokeResult(() => ComObject.Parent)); } }
        public int Count { get { return InvokeResult(() => ComObject.Count); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public VBComponent AddCustom(string progId)
        {
            return new VBComponent(InvokeResult(() => ComObject.AddCustom(progId)));
        }

        public VBComponent AddMTDesigner(int index = 0)
        {
            return new VBComponent(InvokeResult(() => ComObject.AddMTDesigner(index)));
        }

        public VBComponent Item(object index)
        {
            return new VBComponent(InvokeResult(() => ComObject.Item(index)));
        }

        IEnumerator<VBComponent> IEnumerable<VBComponent>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.VBComponents, VBComponent>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.VBComponents, VBComponent>(ComObject);
        }
    }
}