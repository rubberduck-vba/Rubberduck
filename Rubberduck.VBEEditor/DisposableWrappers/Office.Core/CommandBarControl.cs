using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBarControl : SafeComWrapper<Microsoft.Office.Core.CommandBarControl>, IEquatable<CommandBarControl>
    {
        public CommandBarControl(Microsoft.Office.Core.CommandBarControl comObject) 
            : base(comObject)
        {
        }

        public bool BeginsGroup
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.BeginGroup); }
            set { Invoke(() => ComObject.BeginGroup = value); }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.BuiltIn); }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Caption); }
            set { Invoke(() => ComObject.Caption = value); }
        }

        public string DescriptionText
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.DescriptionText); }
            set { Invoke(() => ComObject.DescriptionText = value); }
        }

        public bool IsEnabled
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.Enabled); }
            set { Invoke(() => ComObject.Enabled = value); }
        }

        public int Height
        {
            get { return InvokeResult(() => ComObject.Height); }
            set { Invoke(() => ComObject.Height = value); }
        }

        public int Id
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Id); }
        }

        public int Index
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Index); }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Left); }
        }

        public string OnAction
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.OnAction); }
            set { Invoke(() => ComObject.OnAction = value); }
        }

        public CommandBar Parent
        {
            get { return new CommandBar(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Parent)); }
        }

        public string Parameter
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Parameter); }
            set { Invoke(() => ComObject.Parameter = value); }
        }

        public int Priority
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Priority); }
            set { Invoke(() => ComObject.Priority = value); }
        }

        public string Tag 
        {
            get { return InvokeResult(() => ComObject.Tag); }
            set { Invoke(() => ComObject.Tag = value); }
        }

        public string TooltipText
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.TooltipText); }
            set { Invoke(() => ComObject.TooltipText = value); }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Top); }
        }

        public ControlType Type
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => (ControlType)ComObject.Type); }
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

        public bool IsPriorityDropped
        {
            get { return InvokeResult(() => ComObject.IsPriorityDropped); }
        }

        public void Delete()
        {
            Invoke(() => ComObject.Delete(true));
        }

        public void Execute()
        {
            Invoke(() => ComObject.Execute());
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
                ((int)other.ComObject.Type == (int)Type
                && other.ComObject.Id == Id
                && other.ComObject.Index == Index
                && other.ComObject.BuiltIn == IsBuiltIn
                && ReferenceEquals(other.ComObject.Parent, ComObject.Parent));
        }

        public bool Equals(CommandBarControl other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Office.Core.CommandBarControl>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Type, Id, Index, IsBuiltIn, ComObject.Parent);
        }
    }
}