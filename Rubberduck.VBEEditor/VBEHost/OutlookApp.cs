using System;
using System.Runtime.InteropServices;

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
            throw new NotImplementedException();
        }

        public override void Save()
        {
        }
    }
}