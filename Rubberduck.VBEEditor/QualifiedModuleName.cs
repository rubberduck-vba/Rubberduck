using System;
using Microsoft.Vbe.Interop;
using System.IO;
using System.Linq;

namespace Rubberduck.VBEditor
{
    /// <summary>
    /// Represents a VBComponent or a VBProject.
    /// </summary>
    public struct QualifiedModuleName
    {
        private static string GetDisplayName(VBProject project)
        {

            if (project.Protection == vbext_ProjectProtection.vbext_pp_none)
            {
                //Try reading the top-most document-type component's Properties("Name") value
                //Eg. A Workbook's parent is the application, so read the workbook's name
                try
                {
                    var component = project.VBComponents.Cast<VBComponent>()
                        .FirstOrDefault(comp => comp.Type == vbext_ComponentType.vbext_ct_Document
                                                && comp.Properties.Item("Name").Value != null
                                                && comp.Properties.Item("Parent")
                                                    .Object.Equals(comp.Properties.Item("Application").Object));

                    if (component == null) { return null; }

                    var nameProperty = component.Properties.Cast<Property>().FirstOrDefault(property => property.Name == "Name");
                    return nameProperty == null
                        ? null
                        : nameProperty.Value.ToString();
                }
                catch 
                {
                  //The Properties collection either wasn't available, or didn't have the expected properties
                }

                //Try reading the top-most document-type component's parent's Properties("Name") value
                // Eg. A PowerPoint Slide is top level, but it's parent is a Presentation (that is NOT a vbComponent)
                try
                {
                    var firstOrDefault = project.VBComponents.Cast<VBComponent>()
                        .FirstOrDefault(comp => comp.Type == vbext_ComponentType.vbext_ct_Document
                                                && comp.Properties.Item("Parent").Value != null);
                    if (firstOrDefault != null)
                    {
                        var parentProp = firstOrDefault
                            .Properties.Cast<Property>().FirstOrDefault(property => property.Name == "Parent");

                        Property nameProperty = null;
                        if (parentProp != null && parentProp.Value is Properties)
                        {
                            var props = (Properties)parentProp.Value;
                            nameProperty = props.Cast<Property>().FirstOrDefault(property => property.Name == "Name");
                        }

                        return nameProperty == null
                            ? null
                            : nameProperty.Value.ToString();
                    }
                }
                catch
                {
                    //The Properties collection either wasn't available, or didn't have the expected properties
                }
            }

            //Try reading the filename
            try
            {
                if (!string.IsNullOrEmpty(Path.GetDirectoryName(project.BuildFileName)))
                {
                    return Path.GetFileName(project.FileName);
                }
            }
            catch
            {  //The GetFileName getter probably threw
            }

            return null;
        }

        private static string GetDisplayName(VBComponent component)
        {
            if (component.Type == vbext_ComponentType.vbext_ct_Document)
            {
                //Check for a valid properties collection (some hosts don't validate the Properties method unless the component's designer is open in the host
                try
                {
                    var nameProperty = component.Properties.Item("Name");
                    return nameProperty == null
                        ? null
                        : nameProperty.Value.ToString();
                }
                catch 
                { 
                    //The component isn't open in the host, the Properties Collection is probably inaccessible
                }
            }
            return null;    
        }

        public static string GetProjectId(VBProject project)
        {
            if (project == null)
            {
                return string.Empty;
            }
            return string.IsNullOrEmpty(project.HelpFile) 
                ? project.GetHashCode().ToString() 
                : project.HelpFile;
        }

        public static string GetProjectId(Reference reference)
        {
            var projectName = reference.Name;
            var path = reference.FullPath;
            return new QualifiedModuleName(projectName, path, projectName).ProjectId;
        }

        public QualifiedModuleName(VBProject project)
        {
            _component = null;
            _componentName = null;
            _componentDisplayName = null;
            _project = project;
            _projectName = project.Name;
            _projectPath = string.Empty;
            _projectId = GetProjectId(project);
            _projectDisplayName = GetDisplayName(project);
            _contentHashCode = 0;
        }

        public QualifiedModuleName(VBComponent component)
        {
            _project = null; // field is only assigned when the instance refers to a VBProject.

            _component = component;
            _componentName = component == null ? string.Empty : component.Name;
            _componentDisplayName = GetDisplayName(component);
            _project = component == null ? null : component.Collection.Parent;
            _projectName = _project == null ? string.Empty : _project.Name;
            _projectDisplayName = GetDisplayName(_project);
            _projectPath = string.Empty;
            _projectId = GetProjectId(_project);

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
        /// Creates a QualifiedModuleName for removing renamed declarations.
        /// Do not use this overload.
        /// </summary>
        public QualifiedModuleName(VBComponent component, string oldComponentName)
        {
            _project = null; // field is only assigned when the instance refers to a VBProject.

            _component = component;
            _componentName = oldComponentName;
            _componentDisplayName = GetDisplayName(component);
            _project = component == null ? null : component.Collection.Parent;
            _projectName = _project == null ? string.Empty : _project.Name;
            _projectDisplayName = GetDisplayName(_project);
            _projectPath = string.Empty;
            _projectId = GetProjectId(_project);

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
        public QualifiedModuleName(string projectName, string projectPath, string componentName)
        {
            _project = null;
            _projectName = projectName;
            _projectDisplayName = null;
            _projectPath = projectPath;
            _projectId = (_projectName + ";" + _projectPath).GetHashCode().ToString();
            _componentName = componentName;
            _component = null;
            _componentDisplayName = null;
            _contentHashCode = 0;
        }

        public QualifiedMemberName QualifyMemberName(string member)
        {
            return new QualifiedMemberName(this, member);
        }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }

        private readonly VBProject _project;
        public VBProject Project { get { return _project; } }

        private readonly int _contentHashCode;
        public int ContentHashCode { get { return _contentHashCode; } }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly string _componentName;
        public string ComponentName { get { return _componentName ?? string.Empty; } }

        public string Name { get { return ToString(); } }

        private readonly string _componentDisplayName;
        public string ComponentDisplayName { get {return _componentDisplayName; } }

        private readonly string _projectDisplayName;
        public string ProjectDisplayName { get { return _projectDisplayName; } }


        /// <summary>
        /// returns: "ComponentName (DisplayName)" as typically displayed in VBE Project Explorer
        /// </summary>
        public string ComponentTitle {
            get {
                if (_project != null && _component == null)
                {
                    //handle display of Project component
                    return _projectName + (_projectDisplayName != null ? " (" + _projectDisplayName + ")" : string.Empty);
                }
                else
                {
                    if (_componentDisplayName == _projectDisplayName) 
                    {
                        //handle display of main documents, like ThisWorkbook and ThisDocument
                        return _componentName;
                    }
                    else
                    {
                        //handle display of all other components
                        return _componentName + (_componentDisplayName != null ? " (" +  _componentDisplayName  + ")" : string.Empty);
                    }
                }
            } 
        }

        /// <summary>
        /// returns: "ProjectName (DisplayName)" as typically displayed in VBE Project Explorer
        /// </summary>
        public string ProjectTitle
        {
            get
            {
                return _projectName + (_projectDisplayName != null ? " (" + _projectDisplayName + ")" : string.Empty);
            }
        }

        private readonly string _projectName;

        public string ProjectName
        {
            get
            {
                return _projectName;
            }
        }
        private readonly string _projectPath;

        public string ProjectPath
        {
            get
            {
                return _projectPath;
            }
        }

        public override string ToString()
        {
            return _component == null && string.IsNullOrEmpty(_projectName)
                ? string.Empty
                : (string.IsNullOrEmpty(_projectPath) ? string.Empty : System.IO.Path.GetFileName(_projectPath) + ";")
                     + _projectName + "." + _componentName;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash * 23 + _projectId.GetHashCode();
                hash = hash * 23 + (_componentName ?? string.Empty).GetHashCode();
                return hash;
            }
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
