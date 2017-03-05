using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Rubberduck.VBEditor.Application
{
    [ComVisible(false)]
    public class OutlookApp : HostApplicationBase<Microsoft.Office.Interop.Outlook.Application>
    {
        public OutlookApp() : base("Outlook")
        {
        }

        public override void Run(dynamic declaration)
        {
            // note: does not work. http://stackoverflow.com/q/31954364/1188513
            //var app = Application.GetType();
            //app.InvokeMember(qualifiedMemberName.MemberName, BindingFlags.InvokeMethod | BindingFlags.Default, null, Application, null);
            //Application.Run(qualifiedMemberName.ToString());

            //note: this will work, but not implemented yet http://stackoverflow.com/questions/31954364#36889671
            //TaskItem taskitem = Application.CreateItem(OlItemType.olTaskItem);
            //taskitem.Subject = "Rubberduck";
            //taskitem.Body = qualifiedMemberName.MemberName;

            try
            {
                var app = Application;
                var exp = app.ActiveExplorer();
                CommandBar cb = exp.CommandBars.Add("RubberduckCallbackProxy", Temporary: true);
                CommandBarControl btn = cb.Controls.Add(MsoControlType.msoControlButton, 1);
                btn.OnAction = declaration.QualifiedName.ToString();
                btn.Execute();
                cb.Delete();
            }
            catch {             
            }
        }
    }
}
