using System;
using System.Linq;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Reference : SafeComWrapper<Microsoft.Vbe.Interop.Reference>
    {
        public Reference(Microsoft.Vbe.Interop.Reference comObject) 
            : base(comObject)
        {
        }

        public References Collection { get { return new References(InvokeResult(() => ComObject.Collection)); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public string Name { get { return InvokeResult(() => ComObject.Name); } }
        public Guid Guid { get { return new Guid(InvokeResult(() => ComObject.Guid)); } }
        public int Major { get { return InvokeResult(() => ComObject.Major); } }
        public int Minor { get { return InvokeResult(() => ComObject.Minor); } }
        public string FullPath { get { return InvokeResult(() => ComObject.FullPath); } }
        public bool IsBuiltIn { get { return InvokeResult(() => ComObject.BuiltIn); } }
        public bool IsBroken { get { return InvokeResult(() => ComObject.IsBroken); } }
        public ReferenceKind Type { get { return (ReferenceKind)InvokeResult(() => ComObject.Type); } }
        public string Description { get { return InvokeResult(() => ComObject.Description); } }

        public int Index { get { return Collection.ToList().IndexOf(this); } }
    }
}