using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using MSO = Microsoft.Office.Core;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.Office12
{
    public class CommandBarControl : SafeComWrapper<MSO.CommandBarControl>, ICommandBarControl
    {
        public const bool AddCommandBarControlsTemporarily = false;

        public CommandBarControl(MSO.CommandBarControl target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public bool BeginsGroup
        {
            get => !IsWrappingNullReference && Target.BeginGroup;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.BeginGroup = value;
                }
            }
        }

        public bool IsBuiltIn => !IsWrappingNullReference && Target.BuiltIn;

        public string Caption
        {
            get => IsWrappingNullReference ? string.Empty : Target.Caption;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Caption = CommandBarControlCaptionGuard.ApplyGuard(value);
                }
            }
        }

        public string DescriptionText
        {
            get => IsWrappingNullReference ? string.Empty : Target.DescriptionText;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.DescriptionText = value;
                }
            }
        }

        public bool IsEnabled
        {
            get => !IsWrappingNullReference && Target.Enabled;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Enabled = value;
                }
            }
        }

        public int Height
        {
            get => Target.Height;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Height = value;
                }
            }
        }

        public int Id => IsWrappingNullReference ? 0 : Target.Id;

        public int Index => IsWrappingNullReference ? 0 : Target.Index;

        public int Left => IsWrappingNullReference ? 0 : Target.Left;

        public string OnAction
        {
            get => IsWrappingNullReference ? string.Empty : Target.OnAction;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.OnAction = value;
                }
            }
        }

        public ICommandBar Parent => new CommandBar(IsWrappingNullReference ? null : Target.Parent);

        public string Parameter
        {
            get => IsWrappingNullReference ? string.Empty : Target.Parameter;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Parameter = value;
                }
            }
        }

        public int Priority
        {
            get => IsWrappingNullReference ? 0 : Target.Priority;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Priority = value;
                }
            }
        }

        public string Tag 
        {
            get => Target?.Tag;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Tag = value;
                }
            }
        }

        public string TooltipText
        {
            get => IsWrappingNullReference ? string.Empty : Target.TooltipText;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.TooltipText = value;
                }
            }
        }

        public int Top => IsWrappingNullReference ? 0 : Target.Top;

        public ControlType Type => IsWrappingNullReference ? 0 : (ControlType)Target.Type;

        public bool IsVisible
        {
            get => !IsWrappingNullReference && Target.Visible;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Visible = value;
                }
            }
        }

        public int Width
        {
            get => IsWrappingNullReference ? 0 : Target.Width;
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Width = value;
                }
            }
        }

        public void Delete()
        {
            if (!IsWrappingNullReference)
            {
                Target.Delete(AddCommandBarControlsTemporarily);
            }
        }

        public void Execute()
        {
            if (!IsWrappingNullReference)
            {
                Target.Execute();
            }
        }

        public override bool Equals(ISafeComWrapper<MSO.CommandBarControl> other)
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
            return Equals(other as SafeComWrapper<MSO.CommandBarControl>);
        }

        private int _hashCode;
        public override int GetHashCode()
        {
            if (_hashCode == 0)
            {
                _hashCode = IsWrappingNullReference ? 0 : HashCode.Compute(Type, Id, Index, IsBuiltIn, Target.Parent);
            }

            return _hashCode;
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}