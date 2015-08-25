using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.VBEHost
{
    [ComVisible(false)]
    public class OutlookApp : HostApplicationBase<Microsoft.Office.Interop.Outlook.Application>
    {
        public OutlookApp() : base("Outlook") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var app = Application.GetType();
            app.InvokeMember(qualifiedMemberName.MemberName, BindingFlags.InvokeMethod | BindingFlags.Instance, null, Application, null);
        }
    }
}