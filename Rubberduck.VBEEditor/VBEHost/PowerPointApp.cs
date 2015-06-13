using NetOffice.PowerPointApi;

namespace Rubberduck.VBEditor.VBEHost
{
    public class PowerPointApp : HostApplicationBase<Application>
    {
        public PowerPointApp() : base("PowerPoint") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            object[] paramArray = { }; //PowerPoint requires a paramarray, so we pass it an empty array.

            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call, paramArray);
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            /* Note: Powerpoint supports a `FileName.ppt!Module.method` syntax
             * http://msdn.microsoft.com/en-us/library/office/ff744221(v=office.15).aspx
             */

            return qualifiedMemberName.QualifiedModuleName.Component.Name + "." + qualifiedMemberName.MemberName;

            // todo: verify that the 'FileName.ppt!Module.Method' syntax is real.
            // if a saved presentation can run the above, then the below can just be removed.
            if (!qualifiedMemberName.QualifiedModuleName.Project.Saved)
            {
            }

            var projectFile = qualifiedMemberName.QualifiedModuleName.Project.FileName;
            var moduleName = qualifiedMemberName.QualifiedModuleName.Component.Name;

            return string.Concat(projectFile, "!", moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}