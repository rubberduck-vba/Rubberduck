using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace Rubberduck.VBEditor.VBEHost
{
    public class ExcelApp : HostApplicationBase<Application>
    {
        public ExcelApp() : base("Excel") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            Application.Run(qualifiedMemberName.ToString());
        }

        public override void Save()
        {
            Application.ActiveWorkbook.Save();
        }
    }
}