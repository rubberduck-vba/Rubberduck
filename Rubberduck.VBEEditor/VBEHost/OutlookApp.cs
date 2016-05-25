using System;
using System.Runtime.InteropServices;
//using Microsoft.Office.Interop.Outlook;

namespace Rubberduck.VBEditor.VBEHost
{
    [ComVisible(false)]
    public class OutlookApp : HostApplicationBase<Microsoft.Office.Interop.Outlook.Application>
    {
        public OutlookApp() : base("Outlook")
        {
        }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            // note: does not work. http://stackoverflow.com/q/31954364/1188513
            //var app = Application.GetType();
            //app.InvokeMember(qualifiedMemberName.MemberName, BindingFlags.InvokeMethod | BindingFlags.Default, null, Application, null);
            //Application.Run(qualifiedMemberName.ToString());

            //note: this will work, but not implemented yet http://stackoverflow.com/questions/31954364#36889671
            //TaskItem taskitem = Application.CreateItem(OlItemType.olTaskItem);
            //taskitem.Subject = "Rubberduck";
            //taskitem.Body = qualifiedMemberName.MemberName;

            throw new NotImplementedException();
        }
    }
}
