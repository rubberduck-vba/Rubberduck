using Microsoft.Office.Interop.Word;

namespace Rubberduck.VBEHost
{
    public class WordApp : HostApplicationBase<Application>
    {
        public WordApp() : base("Word") { }

        public override void Run(string target)
        {
            base.Application.Run(target);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}