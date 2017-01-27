using Path = System.IO.Path;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class ExcelApp : HostApplicationBase<Microsoft.Office.Interop.Excel.Application>
    {
        public ExcelApp() : base("Excel") { }
        public ExcelApp(IVBE vbe) : base(vbe, "Excel") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call);
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var module = qualifiedMemberName.QualifiedModuleName;

            var documentName = string.IsNullOrEmpty(module.ProjectPath)
                ? module.ProjectDisplayName
                : Path.GetFileName(module.ProjectPath);

            return string.IsNullOrEmpty(documentName)
                ? qualifiedMemberName.ToString()
                : string.Format("'{0}'!{1}", documentName.Replace("'", "''"), qualifiedMemberName);
        }
    }
}
