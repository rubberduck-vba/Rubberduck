using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Access;

namespace Rubberduck.VBEditor.Application
{
    public class AccessApp : HostApplicationBase<Microsoft.Office.Interop.Access.Application>
    {
        public AccessApp() : base("Access") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call);
        }

        //public List<string> FormDeclarations(QualifiedModuleName qualifiedModuleName)
        //{
        //    Application.DoCmd.OutputTo(AcOutputObjectType.acOutputForm, qualifiedModuleName.Name, AcCommand.acCmdExportText, 
        //        Path.Combine(ExportPath, qualifiedModuleName.Name), null, null, null);
        //}

        //private string ExportPath
        //{
        //    get
        //    {
        //        var assemblyLocation = Assembly.GetAssembly(typeof(AccessApp)).Location;
        //        return Path.GetDirectoryName(assemblyLocation);
        //    }
        //}

        private string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            //Access only supports Project.Procedure syntax. Error occurs if there are naming conflicts.
            // http://msdn.microsoft.com/en-us/library/office/ff193559(v=office.15).aspx
            // https://github.com/retailcoder/Rubberduck/issues/109

            var projectName = qualifiedMemberName.QualifiedModuleName.Project.Name;
            return string.Concat(projectName, ".", qualifiedMemberName.MemberName);
        }
    }
}
