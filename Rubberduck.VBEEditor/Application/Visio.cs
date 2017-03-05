using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class VisioApp : HostApplicationBase<Microsoft.Office.Interop.Visio.Application>
    {
        public VisioApp() : base("Visio") { }
        public VisioApp(IVBE vbe) : base(vbe, "Visio") { }

        public override void Run(dynamic declaration)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Document doc = Application.Documents[declaration.ProjectDisplayName];
                var call = GenerateMethodCall(declaration.QualifiedName);
                doc.ExecuteLine(call);
            }
            catch 
            {
                //Failed to run
            }
        }
        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.ComponentName;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}
