using System;
using System.Reflection;

namespace Rubberduck.VBEditor.VBEHost
{
    public class OutlookApp : HostApplicationBase<Microsoft.Office.Interop.Outlook.Application>
    {
        public OutlookApp() : base("Outlook") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            //Outlook does not support the run method.
            //throw new NotImplementedException("Unit Testing not supported for Outlook");
            var app = Application.GetType();
            app.InvokeMember(qualifiedMemberName.MemberName, BindingFlags.InvokeMethod, null, Application, null);
        }
    }
}