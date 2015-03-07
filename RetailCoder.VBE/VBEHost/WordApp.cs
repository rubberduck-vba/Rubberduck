using Microsoft.Office.Interop.Word;

namespace Rubberduck.VBEHost
{
    public class WordApp : HostApplicationBase<Application>
    {
        public WordApp() : base("Word") { }

        public override void Run(string projectName, string moduleName, string methodName)
        {
            var call = GenerateMethodCall(projectName, moduleName, methodName);
            Application.Run(call);
        }

        protected override string GenerateMethodCall(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}