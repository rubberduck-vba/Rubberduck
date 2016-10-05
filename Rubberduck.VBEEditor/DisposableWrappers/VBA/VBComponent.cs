using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class VBComponent : SafeComWrapper<Microsoft.Vbe.Interop.VBComponent>, IEquatable<VBComponent>
    {
        public VBComponent(Microsoft.Vbe.Interop.VBComponent comObject) 
            : base(comObject)
        {
        }

        public ComponentType Type
        {
            get { return IsWrappingNullReference ? 0 : (ComponentType)InvokeResult(() => ComObject.Type); }
        }

        public CodeModule CodeModule
        {
            get { return new CodeModule(IsWrappingNullReference ? null : InvokeResult(() => ComObject.CodeModule)); }
        }

        public VBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBE)); }
        }

        public VBComponents Collection
        {
            get { return new VBComponents(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Collection)); }
        }

        public Properties Properties
        {
            get { return new Properties(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Properties)); }
        }

        public bool HasOpenDesigner
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.HasOpenDesigner); }
        }

        public string DesignerId
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.DesignerID); }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        // ReSharper disable once ReturnTypeCanBeEnumerable.Global
        public Controls Controls
        {
            get
            {
                var designer = IsWrappingNullReference
                    ? null
                    : InvokeResult(() => ComObject.Designer) as Microsoft.Vbe.Interop.Forms.UserForm;

                return designer == null 
                    ? new Controls(null) 
                    : new Controls(designer.Controls);
            }
        }

        public bool HasDesigner
        {
            get
            {
                if (IsWrappingNullReference)
                {
                    return false;
                }
                var designer = InvokeResult(() => ComObject.Designer);
                var hasDesigner = designer != null;
                return hasDesigner;
            }
        }

        public Window DesignerWindow()
        {
            return new Window(IsWrappingNullReference ? null : InvokeResult(() => ComObject.DesignerWindow()));
        }

        public void Activate()
        {
            Invoke(() => ComObject.Activate());
        }

        public bool IsSaved { get { return !IsWrappingNullReference && InvokeResult(() => ComObject.Saved); } }

        public void Export(string path)
        {
            Invoke(() => ComObject.Export(path));
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                DesignerWindow().Release();
                Controls.Release();
                Properties.Release();
                CodeModule.Release();
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBComponent> other)
        {
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
        }

        public bool Equals(VBComponent other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBComponent>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}