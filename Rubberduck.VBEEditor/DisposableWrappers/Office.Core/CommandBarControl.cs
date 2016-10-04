namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBarControl : SafeComWrapper<Microsoft.Office.Core.CommandBarControl>
    {
        public CommandBarControl(Microsoft.Office.Core.CommandBarControl comObject) 
            : base(comObject)
        {
        }

        public void Delete()
        {
            Invoke(() => ComObject.Delete(true));
        }

        public void Execute()
        {
            Invoke(() => ComObject.Execute());
        }

        public bool BeginsGroup
        {
            get { return InvokeResult(() => ComObject.BeginGroup); }
            set { Invoke(() => ComObject.BeginGroup = value); }
        }

        public bool IsBuiltIn { get { return InvokeResult(() => ComObject.BuiltIn); } }

        public string Caption
        {
            get { return InvokeResult(() => ComObject.Caption); }
            set { Invoke(() => ComObject.Caption = value); }
        }

        public string DescriptionText
        {
            get { return InvokeResult(() => ComObject.DescriptionText); }
            set { Invoke(() => ComObject.DescriptionText = value); }
        }

        public bool IsEnabled
        {
            get { return InvokeResult(() => ComObject.Enabled); }
            set { Invoke(() => ComObject.Enabled = value); }
        }

        public int Height { get; set; }
        public int Id { get { return InvokeResult(() => ComObject.Id); } }
        public int Index { get { return InvokeResult(() => ComObject.Index); } }
        public int Left { get { return InvokeResult(() => ComObject.Left); } }

        public string OnAction
        {
            get { return InvokeResult(() => ComObject.OnAction); }
            set { Invoke(() => ComObject.OnAction = value); }
        }

        public CommandBar Parent { get { return new CommandBar(InvokeResult(() => ComObject.Parent)); } }

        public string Parameter
        {
            get { return InvokeResult(() => ComObject.Parameter); }
            set { Invoke(() => ComObject.Parameter = value); }
        }

        public int Priority
        {
            get { return InvokeResult(() => ComObject.Priority); }
            set { Invoke(() => ComObject.Priority = value); }
        }

        public string Tag 
        {
            get { return InvokeResult(() => ComObject.Tag); }
            set { Invoke(() => ComObject.Tag = value); }
        }

        public string TooltipText
        {
            get { return InvokeResult(() => ComObject.TooltipText); }
            set { Invoke(() => ComObject.TooltipText = value); }
        }

        public int Top { get { return InvokeResult(() => ComObject.Top); } }

        public ControlType Type { get { return InvokeResult(() => (ControlType)ComObject.Type); } }

        public bool IsVisible
        {
            get { return InvokeResult(() => ComObject.Visible); }
            set { Invoke(() => ComObject.Visible = value); }
        }

        public int Width
        {
            get { return InvokeResult(() => ComObject.Width); }
            set { Invoke(() => ComObject.Width = value); }
        }

        public bool IsPriorityDropped { get { return InvokeResult(() => ComObject.IsPriorityDropped); } }
    }
}