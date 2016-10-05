using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBar : SafeComWrapper<Microsoft.Office.Core.CommandBar>, IEquatable<CommandBar>
    {
        public CommandBar(Microsoft.Office.Core.CommandBar comObject) 
            : base(comObject)
        {
        }

        public int Id
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Id); }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.BuiltIn); }
        }

        public CommandBarControls Controls
        {
            get { return new CommandBarControls(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Controls)); }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.Enabled); }
        }

        public int Height
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Height); }
            set { Invoke(() => ComObject.Height = value); }
        }

        public int Index
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Index); }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Left); }
            set { Invoke(() => ComObject.Left = value); }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public CommandBarPosition Position
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => (CommandBarPosition)ComObject.Position); }
            set { Invoke(() => ComObject.Position = (Microsoft.Office.Core.MsoBarPosition)value); }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Top); }
            set { Invoke(() => ComObject.Top = value); }
        }

        public CommandBarType Type
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => (CommandBarType)ComObject.Type); }
        }

        public bool IsVisible
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.Visible); }
            set { Invoke(() => ComObject.Visible = value); }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Width); }
            set { Invoke(() => ComObject.Width = value); }
        }

        public CommandBarControl FindControl(int id)
        {
            return new CommandBarControl(InvokeResult(() => ComObject.FindControl(Id: id)));
        }

        public CommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(InvokeResult(() => ComObject.FindControl(type, id)));
        }
        
        public void Delete()
        {
            Invoke(() => ComObject.Delete());
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Controls.Release();
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Office.Core.CommandBar> other)
        {
            return IsEqualIfNull(other) || 
                (other != null
                && (int)other.ComObject.Type == (int)Type 
                && other.ComObject.Id == Id 
                && other.ComObject.Index == Index
                && other.ComObject.BuiltIn == IsBuiltIn 
                && ReferenceEquals(other.ComObject.Parent, ComObject.Parent));
        }

        public bool Equals(CommandBar other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBar>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Type, Id, Index, IsBuiltIn, ComObject.Parent);
        }
    }
}
