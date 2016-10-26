using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class ExcelApp : HostApplicationBase<Microsoft.Office.Interop.Excel.Application>
    {
        public ExcelApp() : base("Excel") { }
        public ExcelApp(IVBE vbe) : base(vbe, "Excel") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            string call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call);
        }

       protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var documentName = qualifiedMemberName.QualifiedModuleName.ProjectDisplayName;
            return string.Concat(documentName, "!", qualifiedMemberName.ToString());
        }
    }
}
