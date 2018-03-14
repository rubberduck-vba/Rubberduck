using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class PowerPointApp : HostApplicationBase<Microsoft.Office.Interop.PowerPoint.Application>
    {
        private readonly IVBE _vbe;

        public PowerPointApp(IVBE vbe) : base(vbe, "PowerPoint")
        {
            _vbe = vbe;
        }

        public override void Run(dynamic declaration)
        {
            var methodCall = GenerateMethodCall(declaration.QualifiedName);
            if (methodCall == null)
            {
                return;
            }

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
                using (var activeProject = _vbe.ActiveVBProject)
                {
                    if (activeProject.IsWrappingNullReference)
                    {
                        return null; // if there's no active project, we can't generate the call
                    }
                }
                return $"{qualifiedMemberName.QualifiedModuleName.ComponentName}.{qualifiedMemberName.MemberName}";
            }

            var moduleName = qualifiedMemberName.QualifiedModuleName.ComponentName;
            return $"{path}!{moduleName}.{qualifiedMemberName.MemberName}";
        }
    }
}
