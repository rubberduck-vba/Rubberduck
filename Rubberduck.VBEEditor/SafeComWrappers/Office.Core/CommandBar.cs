using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBar : SafeComWrapper<Microsoft.Office.Core.CommandBar>, ICommandBar
    {
        public CommandBar(Microsoft.Office.Core.CommandBar target) 
            : base(target)
        {
        }

        public int Id
        {
            get { return IsWrappingNullReference ? 0 : Target.Id; }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && Target.BuiltIn; }
        }

        public ICommandBarControls Controls
        {
            get { return new CommandBarControls(IsWrappingNullReference ? null : Target.Controls); }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && Target.Enabled; }
            set { if (!IsWrappingNullReference) Target.Enabled = value; }
        }

        public int Height
        {
            get { return IsWrappingNullReference ? 0 : Target.Height; }
            set { if (!IsWrappingNullReference) Target.Height = value; }
        }

        public int Index
        {
            get { return IsWrappingNullReference ? 0 : Target.Index; }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : Target.Left; }
            set { if (!IsWrappingNullReference) Target.Left = value; }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        public CommandBarPosition Position
        {
            get { return IsWrappingNullReference ? 0 : (CommandBarPosition)Target.Position; }
            set { if (!IsWrappingNullReference) Target.Position = (Microsoft.Office.Core.MsoBarPosition)value; }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : Target.Top; }
            set { if (!IsWrappingNullReference) Target.Top = value; }
        }

        public CommandBarType Type
        {
            get { return IsWrappingNullReference ? 0 : (CommandBarType)Target.Type; }
        }

        public bool IsVisible
        {
            get { return !IsWrappingNullReference && Target.Visible; }
            set { if (!IsWrappingNullReference) Target.Visible = value; }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : Target.Width; }
            set { if (!IsWrappingNullReference) Target.Width = value; }
        }

        public ICommandBarControl FindControl(int id)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.FindControl(Id: id));
        }

        public ICommandBarControl FindControl(ControlType type, int id)
        {
            return new CommandBarControl(IsWrappingNullReference ? null : Target.FindControl(type, id));
        }
        
        public void Delete()
        {
            if (!IsWrappingNullReference) Target.Delete();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        Controls.Release();
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(ISafeComWrapper<Microsoft.Office.Core.CommandBar> other)
        {
            return IsEqualIfNull(other) || 
                (other != null
                && (int)other.Target.Type == (int)Type 
                && other.Target.Id == Id 
                && other.Target.Index == Index
                && other.Target.BuiltIn == IsBuiltIn 
                && ReferenceEquals(other.Target.Parent, Target.Parent));
        }

        public bool Equals(ICommandBar other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBar>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Type, Id, Index, IsBuiltIn, Target.Parent);
        }
    }
}
