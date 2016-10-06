using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarControl : SafeComWrapper<Microsoft.Office.Core.CommandBarControl>, ICommandBarControl
    {
        public CommandBarControl(Microsoft.Office.Core.CommandBarControl comObject) 
            : base(comObject)
        {
        }

        public bool BeginsGroup
        {
            get { return !IsWrappingNullReference && ComObject.BeginGroup; }
            set { ComObject.BeginGroup = value; }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && ComObject.BuiltIn; }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Caption; }
            set { ComObject.Caption = value; }
        }

        public string DescriptionText
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.DescriptionText; }
            set { ComObject.DescriptionText = value; }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && ComObject.Enabled; }
            set { ComObject.Enabled = value; }
        }

        public int Height
        {
            get { return ComObject.Height; }
            set { ComObject.Height = value; }
        }

        public int Id
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Id; }
        }

        public int Index
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Index; }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Left; }
        }

        public string OnAction
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.OnAction; }
            set { ComObject.OnAction = value; }
        }

        public ICommandBar Parent
        {
            get { return new CommandBar(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public string Parameter
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Parameter; }
            set { ComObject.Parameter = value; }
        }

        public int Priority
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Priority; }
            set { ComObject.Priority = value; }
        }

        public string Tag 
        {
            get { return ComObject.Tag; }
            set { ComObject.Tag = value; }
        }

        public string TooltipText
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.TooltipText; }
            set { ComObject.TooltipText = value; }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Top; }
        }

        public ControlType Type
        {
            get { return IsWrappingNullReference ? 0 : (ControlType)ComObject.Type; }
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

        public bool IsPriorityDropped
        {
            get { return ComObject.IsPriorityDropped; }
        }

        public void Delete()
        {
            ComObject.Delete(true);
        }

        public void Execute()
        {
            ComObject.Execute();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Office.Core.CommandBarControl> other)
        {
            return IsEqualIfNull(other) ||
                (other != null 
                && (int)other.ComObject.Type == (int)Type
                && other.ComObject.Id == Id
                && other.ComObject.Index == Index
                && other.ComObject.BuiltIn == IsBuiltIn
                && ReferenceEquals(other.ComObject.Parent, ComObject.Parent));
        }

        public bool Equals(ICommandBarControl other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBarControl>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Type, Id, Index, IsBuiltIn, ComObject.Parent);
        }
    }
}