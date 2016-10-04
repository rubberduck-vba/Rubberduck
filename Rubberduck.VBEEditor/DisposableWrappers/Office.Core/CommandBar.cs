namespace Rubberduck.VBEditor.DisposableWrappers.Office.Core
{
    public class CommandBar : SafeComWrapper<Microsoft.Office.Core.CommandBar>
    {
        public CommandBar(Microsoft.Office.Core.CommandBar comObject) 
            : base(comObject)
        {
        }

        public void Delete()
        {
            Invoke(() => ComObject.Delete());
        }

        public bool IsBuiltIn { get { return InvokeResult(() => ComObject.BuiltIn); } }
        public CommandBarControls Controls { get { return new CommandBarControls(InvokeResult(() => ComObject.Controls));} }
        public bool IsEnabled { get { return InvokeResult(() => ComObject.Enabled); } }

        public int Height
        {
            get { return InvokeResult(() => ComObject.Height); }
            set { Invoke(() => ComObject.Height = value); }
        }

        public int Index { get { return InvokeResult(() => ComObject.Index); } }

        public int Left
        {
            get { return InvokeResult(() => ComObject.Left); }
            set { Invoke(() => ComObject.Left = value); }
        }

        public string Name
        {
            get { return InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public CommandBarPosition Position
        {
            get { return InvokeResult(() => (CommandBarPosition)ComObject.Position); }
            set { Invoke(() => ComObject.Position = (Microsoft.Office.Core.MsoBarPosition)value); }
        }

        public int Top
        {
            get { return InvokeResult(() => ComObject.Top); }
            set { Invoke(() => ComObject.Top = value); }
        }
        public CommandBarType Type { get { return InvokeResult(() => (CommandBarType)ComObject.Type); } }

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

        public int Id { get { return InvokeResult(() => ComObject.Id); } }
    }
}
