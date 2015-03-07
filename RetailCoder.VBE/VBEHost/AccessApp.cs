using Microsoft.Office.Interop.Access;

namespace Rubberduck.VBEHost
{
    public class AccessApp : HostApplicationBase<Application>
    {
        public AccessApp() : base("Access") { }

        public override void Run(string projectName, string moduleName, string methodName)
        {
            var call = GenerateMethodCall(projectName, moduleName, methodName);
            Application.Run(call);
        }

        protected override string GenerateMethodCall(string projectName, string moduleName, string methodName)
        {
            //Access only supports Project.Procedure syntax. Error occurs if there are naming conflicts.
            // http://msdn.microsoft.com/en-us/library/office/ff193559(v=office.15).aspx
            // https://github.com/retailcoder/Rubberduck/issues/109

            return string.Concat(projectName, ".", methodName);
        }
    }
}