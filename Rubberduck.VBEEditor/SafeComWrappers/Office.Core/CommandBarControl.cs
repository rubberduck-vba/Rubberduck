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
            set { if (!IsWrappingNullReference) Target.BeginGroup = value; }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && Target.BuiltIn; }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Caption; }
            set { if (!IsWrappingNullReference) Target.Caption = value; }
        }

        public string DescriptionText
        {
            get { return IsWrappingNullReference ? string.Empty : Target.DescriptionText; }
            set { if (!IsWrappingNullReference) Target.DescriptionText = value; }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && Target.Enabled; }
            set { if (!IsWrappingNullReference) Target.Enabled = value; }
        }

        public int Height
        {
            get { return Target.Height; }
            set { if (!IsWrappingNullReference) Target.Height = value; }
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
            set { if (!IsWrappingNullReference) Target.OnAction = value; }
        }

        public ICommandBar Parent
        {
            get { return new CommandBar(IsWrappingNullReference ? null : Target.Parent); }
        }

        public string Parameter
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Parameter; }
            set { if (!IsWrappingNullReference) Target.Parameter = value; }
        }

        public int Priority
        {
            get { return IsWrappingNullReference ? 0 : Target.Priority; }
            set { if (!IsWrappingNullReference) Target.Priority = value; }
        }

        public string Tag 
        {
            get { return Target.Tag; }
            set { if (!IsWrappingNullReference) Target.Tag = value; }
        }

        public string TooltipText
        {
            get { return IsWrappingNullReference ? string.Empty : Target.TooltipText; }
            set { if (!IsWrappingNullReference) Target.TooltipText = value; }
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
            set { if (!IsWrappingNullReference) Target.Visible = value; }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : Target.Width; }
            set { if (!IsWrappingNullReference) Target.Width = value; }
        }

        public bool IsPriorityDropped
        {
            get { return (!IsWrappingNullReference) && Target.IsPriorityDropped; }
        }

        public void Delete()
        {
            if (!IsWrappingNullReference) Target.Delete(true);
        }

        public void Execute()
        {
            if (!IsWrappingNullReference) Target.Execute();
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