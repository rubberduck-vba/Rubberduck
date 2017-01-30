using Path = System.IO.Path;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class ExcelApp : HostApplicationBase<Microsoft.Office.Interop.Excel.Application>
    {
        public ExcelApp() : base("Excel") { }
        public ExcelApp(IVBE vbe) : base(vbe, "Excel") { }

        public override void Run(dynamic declaration)
        {
            var call = GenerateMethodCall(declaration);
            Application.Run(call);
        }

        protected virtual string GenerateMethodCall(dynamic declaration)
        {
            var qualifiedMemberName = declaration.QualifiedName;
            var module = qualifiedMemberName.QualifiedModuleName;

            var documentName = string.IsNullOrEmpty(module.ProjectPath)
                ? declaration.ProjectDisplayName
                : Path.GetFileName(module.ProjectPath);

            return string.IsNullOrEmpty(documentName)
                ? qualifiedMemberName.ToString()
                : string.Format("'{0}'!{1}", documentName.Replace("'", "''"), qualifiedMemberName);
        }
    }
}
