using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor
{
    /// <summary>
    /// Represents a VBComponent or a VBProject.
    /// </summary>
    public struct QualifiedModuleName
    {
        public QualifiedModuleName(VBProject project)  
        {
            _component = null;
            _componentName = null;
            _project = project;
            _projectName = project.Name;
            _projectHashCode = project.GetHashCode();
            _contentHashCode = 0;  
       }

        public QualifiedModuleName(VBComponent component)
        {
            _project = null; // field is only assigned when the instance refers to a VBProject.

            _component = component;
            _componentName = component == null ? string.Empty : component.Name;
            _projectName = component == null ? string.Empty : component.Collection.Parent.Name;
            _projectHashCode = component == null ? 0 : component.Collection.Parent.GetHashCode();

            _contentHashCode = 0;
            if (component == null)
            {
                return;
            }

            var module = component.CodeModule;
            _contentHashCode = module.CountOfLines > 0
                // ReSharper disable once UseIndexedProperty
                ? module.get_Lines(1, module.CountOfLines).GetHashCode() 
                : 0;
        }

        /// <summary>
        /// Creates a QualifiedModuleName for a built-in declaration.
        /// Do not use this overload for user declarations.
        /// </summary>
        public QualifiedModuleName(string projectName, string componentName)
        {
            _project = null; // field is only assigned when the instance refers to a VBProject.

            _projectName = projectName;
            _componentName = componentName;
            _component = null;
            _contentHashCode = componentName.GetHashCode();
            _projectHashCode = projectName.GetHashCode();
        }

        public QualifiedMemberName QualifyMemberName(string member)
        {
            return new QualifiedMemberName(this, member);
        }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }

        private readonly VBProject _project;
        public VBProject Project { get { return _project ?? (_component == null ? null : _component.Collection.Parent); } }

        private readonly int _projectHashCode;
        public int ProjectHashCode { get { return _projectHashCode; } }

        private readonly int _contentHashCode;

        private readonly string _projectName;
        public string ProjectName { get { return _projectName;} }

        private readonly string _componentName;
        public string ComponentName { get { return _componentName; } }

        public string Name { get { return ToString(); } }

        public override string ToString()
        {
            return _component == null && string.IsNullOrEmpty(_projectName) ? string.Empty : _projectName + "." + _componentName;
        }

        public override int GetHashCode()
        {
            return _component == null ? 0 : _component.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj == null) { return false; }

            try
            {
                var other = (QualifiedModuleName)obj;
                if (other.Component == null)
                {
                    return other.ProjectName == ProjectName && other.ComponentName == ComponentName;
                }

                var result = other.Project == Project 
                    && other.ProjectName == ProjectName
                    && other.ComponentName == ComponentName 
                    && other._contentHashCode == _contentHashCode;
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