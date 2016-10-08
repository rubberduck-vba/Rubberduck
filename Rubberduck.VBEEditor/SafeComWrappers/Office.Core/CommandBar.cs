using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBar : SafeComWrapper<Microsoft.Office.Core.CommandBar>, ICommandBar
    {
        public CommandBar(Microsoft.Office.Core.CommandBar comObject) 
            : base(comObject)
        {
        }

        public int Id
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Id; }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && ComObject.BuiltIn; }
        }

        public ICommandBarControls Controls
        {
            get { return new CommandBarControls(IsWrappingNullReference ? null : ComObject.Controls); }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && ComObject.Enabled; }
            set { ComObject.Enabled = value; }
        }

        public int Height
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Height; }
            set { ComObject.Height = value; }
        }

        public int Index
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Index; }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Left; }
            set { ComObject.Left = value; }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Name; }
            set { ComObject.Name = value; }
        }

        public CommandBarPosition Position
        {
            get { return IsWrappingNullReference ? 0 : (CommandBarPosition)ComObject.Position; }
            set { ComObject.Position = (Microsoft.Office.Core.MsoBarPosition)value; }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Top; }
            set { ComObject.Top = value; }
        }

        public CommandBarType Type
        {
            get { return IsWrappingNullReference ? 0 : (CommandBarType)ComObject.Type; }
        }

        public bool IsVisible
        {
            get { return !IsWrappingNullReference && ComObject.Visible; }
            set { ComObject.Visible = value; }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Width; }
            set { ComObject.Width = value; }
        }

        public ICommandBarControl FindControl(int id)
        {
            return new CommandBarControl(ComObject.FindControl(Id: id));
        }

        public ICommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(ComObject.FindControl(type, id));
        }
        
        public void Delete()
        {
            ComObject.Delete();
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

        public bool Equals(ICommandBar other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBar>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Type, Id, Index, IsBuiltIn, ComObject.Parent);
        }
    }
}
