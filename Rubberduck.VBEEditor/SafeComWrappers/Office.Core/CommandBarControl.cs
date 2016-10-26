using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarControl : SafeComWrapper<Microsoft.Office.Core.CommandBarControl>, ICommandBarControl
    {
        public CommandBarControl(Microsoft.Office.Core.CommandBarControl target) 
            : base(target)
        {
        }

        public bool BeginsGroup
        {
            get { return !IsWrappingNullReference && Target.BeginGroup; }
            set { Target.BeginGroup = value; }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && Target.BuiltIn; }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Caption; }
            set { Target.Caption = value; }
        }

        public string DescriptionText
        {
            get { return IsWrappingNullReference ? string.Empty : Target.DescriptionText; }
            set { Target.DescriptionText = value; }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && Target.Enabled; }
            set { Target.Enabled = value; }
        }

        public int Height
        {
            get { return Target.Height; }
            set { Target.Height = value; }
        }

        public int Id
        {
            get { return IsWrappingNullReference ? 0 : Target.Id; }
        }

        public int Index
        {
            get { return IsWrappingNullReference ? 0 : Target.Index; }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : Target.Left; }
        }

        public string OnAction
        {
            get { return IsWrappingNullReference ? string.Empty : Target.OnAction; }
            set { Target.OnAction = value; }
        }

        public ICommandBar Parent
        {
            get { return new CommandBar(IsWrappingNullReference ? null : Target.Parent); }
        }

        public string Parameter
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Parameter; }
            set { Target.Parameter = value; }
        }

        public int Priority
        {
            get { return IsWrappingNullReference ? 0 : Target.Priority; }
            set { Target.Priority = value; }
        }

        public string Tag 
        {
            get { return Target.Tag; }
            set { Target.Tag = value; }
        }

        public string TooltipText
        {
            get { return IsWrappingNullReference ? string.Empty : Target.TooltipText; }
            set { Target.TooltipText = value; }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : Target.Top; }
        }

        public ControlType Type
        {
            get { return IsWrappingNullReference ? 0 : (ControlType)Target.Type; }
        }

        public bool IsVisible
        {
            get { return !IsWrappingNullReference && Target.Visible; }
            set { Target.Visible = value; }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : Target.Width; }
            set { Target.Width = value; }
        }

        public bool IsPriorityDropped
        {
            get { return Target.IsPriorityDropped; }
        }

        public void Delete()
        {
            Target.Delete(true);
        }

        public void Execute()
        {
            Target.Execute();
        }

        public override bool Equals(ISafeComWrapper<Microsoft.Office.Core.CommandBarControl> other)
        {
            return IsEqualIfNull(other) ||
                (other != null 
                && (int)other.Target.Type == (int)Type
                && other.Target.Id == Id
                && other.Target.Index == Index
                && other.Target.BuiltIn == IsBuiltIn
                && ReferenceEquals(other.Target.Parent, Target.Parent));
        }

        public bool Equals(ICommandBarControl other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBarControl>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Type, Id, Index, IsBuiltIn, Target.Parent);
        }
    }
}