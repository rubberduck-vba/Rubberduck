
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public struct QualifiedModuleName
    {
        public QualifiedModuleName(string projectName, string module, VBProject project, int contentHashCode)
        {
            _project = project;
            _contentHashCode = contentHashCode;
            _projectName = projectName;
            _module = module;
        }

        public static QualifiedModuleName Empty { get { return new QualifiedModuleName(string.Empty, string.Empty, null, default(int)); } }

        private readonly VBProject _project;
        public VBProject Project { get { return _project; } }

        private readonly int _contentHashCode;
        public int ContentHashCode { get { return _contentHashCode; } }

        private readonly string _projectName;
        public string ProjectName { get { return _projectName; } }

        private readonly string _module;
        public string ModuleName { get { return _module; } }

        public override string ToString()
        {
            return _projectName + "." + _module;
        }

        public override int GetHashCode()
        {
            return (_project.ToString() + _contentHashCode.ToString() + ToString()).GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var other = (QualifiedModuleName)obj;

            return other.ProjectName == ProjectName
                   && other.ModuleName == ModuleName
                   && other.ContentHashCode == ContentHashCode;
        }

        public static bool operator ==(QualifiedModuleName a, QualifiedModuleName b)
        {
            return a.Equals(b);
        }

        public static bool operator !=(QualifiedModuleName a, QualifiedModuleName b)
        {
            return !a.Equals(b);
        }
    }
}