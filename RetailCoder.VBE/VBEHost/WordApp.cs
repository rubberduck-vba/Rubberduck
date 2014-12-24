using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace Rubberduck
{
    [ComVisible(false)]
    public class WordApp : HostApplicationBase<Application>
    {
        public WordApp() : base("Word") { }

        public override void Run(string target)
        {
            base._application.Run(target);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}