using System;
using System.Globalization;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    /// <summary>
    /// Represents a VBComponent or a VBProject.
    /// </summary>
    public struct QualifiedModuleName
    {
        public static string GetProjectId(IVBProject project)
        {
            if (project.IsWrappingNullReference)
            {
                return string.Empty;
            }

            if (string.IsNullOrEmpty(project.HelpFile))
            {
                project.HelpFile = project.GetHashCode().ToString(CultureInfo.InvariantCulture);
            }

            return project.HelpFile;
        }

        public static string GetProjectId(IReference reference)
        {
            var projectName = reference.Name;
            return new QualifiedModuleName(projectName, reference.FullPath, projectName).ProjectId;
        }

        public QualifiedModuleName(IVBProject project)
        {
            _component = null;
            _componentName = null;
            _componentType = ComponentType.Undefined;
            _projectName = project.Name;
            _projectPath = string.Empty;
            _projectId = GetProjectId(project);           
            _contentHashCode = 0;
        }

        public QualifiedModuleName(IVBComponent component)
        {
            _componentType = component.Type;
            _component = component;
            _componentName = component.IsWrappingNullReference ? string.Empty : component.Name;

            _contentHashCode = 0;
            if (!component.IsWrappingNullReference)
            {
                var module = _component.CodeModule;
                _contentHashCode = module.CountOfLines > 0
                    ? module.GetLines(1, module.CountOfLines).GetHashCode()
                    : 0;
            }

            var project = component.Collection.Parent;
            _projectName = project == null ? string.Empty : project.Name;
            _projectPath = string.Empty;
            _projectId = GetProjectId(project);
        }

        /// <summary>
        /// Creates a QualifiedModuleName for a built-in declaration.
        /// Do not use this overload for user declarations.
        /// </summary>
        public QualifiedModuleName(string projectName, string projectPath, string componentName)
        {
            _projectName = projectName;
            _projectPath = projectPath;
            _projectId = string.Format("{0};{1}", _projectName, _projectPath).GetHashCode().ToString(CultureInfo.InvariantCulture);
            _componentName = componentName;
            _component = null;
            _componentType = ComponentType.ComComponent;
            _contentHashCode = 0;
        }

        public QualifiedMemberName QualifyMemberName(string member)
        {
            return new QualifiedMemberName(this, member);
        }

        private readonly IVBComponent _component;
        public IVBComponent Component { get { return _component; } }

        private readonly ComponentType _componentType;
        public ComponentType ComponentType { get { return _componentType; } }

        private readonly int _contentHashCode;
        public int ContentHashCode { get { return _contentHashCode; } }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly string _componentName;
        public string ComponentName { get { return _componentName ?? string.Empty; } }

        public string Name { get { return ToString(); } }

        private readonly string _projectName;
        public string ProjectName { get { return _projectName ?? string.Empty; } }

        private readonly string _projectPath;
        public string ProjectPath { get { return _projectPath; } }

        public override string ToString()
        {
            return string.IsNullOrEmpty(_componentName) && string.IsNullOrEmpty(_projectName)
                ? string.Empty
                : (string.IsNullOrEmpty(_projectPath) ? string.Empty : System.IO.Path.GetFileName(_projectPath) + ";")
                     + _projectName + "." + _componentName;
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(_projectId, _componentName ?? string.Empty);
        }

        public override bool Equals(object obj)
        {
            if (obj == null) { return false; }

            try
            {
                var other = (QualifiedModuleName)obj;
                var result = other.ProjectId == ProjectId && other.ComponentName == ComponentName;
                return result;
            }
            catch (InvalidCastException)
            {
                return false;
            }
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
