using Microsoft.Office.Interop.Word;
using Rubberduck.Parsing;

namespace Rubberduck.VBEHost
{
    public class WordApp : HostApplicationBase<Application>
    {
        public WordApp() : base("Word") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call);
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.Component.Name;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}