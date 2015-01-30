using System.Runtime.InteropServices;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public struct QualifiedModuleName
    {
        public QualifiedModuleName(string project, string module)
        {
            _project = project;
            _module = module;
        }

        private readonly string _project;
        public string ProjectName { get { return _project; } }

        private readonly string _module;
        public string ModuleName { get { return _module; } }
    }
}