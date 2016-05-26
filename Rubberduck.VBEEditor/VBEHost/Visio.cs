using Microsoft.Office.Interop.Visio;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEHost
{
    public class VisioApp : HostApplicationBase<Microsoft.Office.Interop.Visio.Application>
    {
        public VisioApp() : base("Visio") { }
        public VisioApp(VBE vbe) : base(vbe, "Visio") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            try
            {
                Document doc = Application.Documents[qualifiedMemberName.QualifiedModuleName.ProjectDisplayName];
                var call = GenerateMethodCall(qualifiedMemberName);
                doc.ExecuteLine(call);
            }
            catch 
            {
                //Failed to run
            }
        }
        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.Component.Name;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}
