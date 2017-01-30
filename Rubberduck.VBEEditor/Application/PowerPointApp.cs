namespace Rubberduck.VBEditor.Application
{
    public class PowerPointApp : HostApplicationBase<Microsoft.Office.Interop.PowerPoint.Application>
    {
        public PowerPointApp() : base("PowerPoint") { }

        public override void Run(dynamic declaration)
        {
            var methodCall = GenerateMethodCall(declaration.QualifiedName);
            if (methodCall == null) { return; }

            //PowerPoint requires a paramarray, so we pass it an empty array:
            object[] paramArray = { };
            Application.Run(methodCall, ref paramArray);
        }

        private string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            /* Note: Powerpoint supports a `FileName.ppt!Module.method` syntax
            ** http://msdn.microsoft.com/en-us/library/office/ff744221(v=office.15).aspx
            */
            var path = qualifiedMemberName.QualifiedModuleName.ProjectPath;
            if (string.IsNullOrEmpty(path))
            {
                // if project isn't saved yet, we can't qualify the method call: this only works with the active project.
                return qualifiedMemberName.QualifiedModuleName.Component.VBE.ActiveVBProject.IsWrappingNullReference
                    ? null // if there's no active project, we can't generate the call
                    : qualifiedMemberName.QualifiedModuleName.ComponentName + "." + qualifiedMemberName.MemberName;
            }

            var moduleName = qualifiedMemberName.QualifiedModuleName.ComponentName;
            return string.Concat(path, "!", moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}
