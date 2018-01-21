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
            Component = null;
            _componentName = null;
            ComponentType = ComponentType.Undefined;
            _projectName = project.Name;
            ProjectPath = string.Empty;
            ProjectId = GetProjectId(project);           
            ContentHashCode = 0;
        }

        public QualifiedModuleName(IVBComponent component)
        {
            ComponentType = component.Type;
            Component = component;
            _componentName = component.IsWrappingNullReference ? string.Empty : component.Name;

            ContentHashCode = 0;
            if (!Component.IsWrappingNullReference)
            {
                using (var module = Component.CodeModule)
                {
                    ContentHashCode = module.CountOfLines > 0
                        ? module.GetLines(1, module.CountOfLines).GetHashCode()
                        : 0;
                }
            }

            IVBProject project;
            using (var components = component.Collection)
            {
                project = components.Parent;
            }
            _projectName = project == null ? string.Empty : project.Name;
            ProjectPath = string.Empty;
            ProjectId = GetProjectId(project);
        }

        /// <summary>
        /// Creates a QualifiedModuleName for a built-in declaration.
        /// Do not use this overload for user declarations.
        /// </summary>
        public QualifiedModuleName(string projectName, string projectPath, string componentName)
        {
            _projectName = projectName;
            ProjectPath = projectPath;
            ProjectId = $"{_projectName};{ProjectPath}".GetHashCode().ToString(CultureInfo.InvariantCulture);
            _componentName = componentName;
            Component = null;
            ComponentType = ComponentType.ComComponent;
            ContentHashCode = 0;
        }

        public QualifiedMemberName QualifyMemberName(string member)
        {
            return new QualifiedMemberName(this, member);
        }

        public IVBComponent Component { get; }

        public ComponentType ComponentType { get; }

        public int ContentHashCode { get; }

        public string ProjectId { get; }

        private readonly string _componentName;
        public string ComponentName => _componentName ?? string.Empty;

        public string Name => ToString();

        private readonly string _projectName;
        public string ProjectName => _projectName ?? string.Empty;

        public string ProjectPath { get; }

        public override string ToString()
        {
            return string.IsNullOrEmpty(_componentName) && string.IsNullOrEmpty(_projectName)
                ? string.Empty
                : (string.IsNullOrEmpty(ProjectPath) ? string.Empty : System.IO.Path.GetFileName(ProjectPath) + ";")
                     + $"{_projectName}.{_componentName}";
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(ProjectId ?? string.Empty, _componentName ?? string.Empty);
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
