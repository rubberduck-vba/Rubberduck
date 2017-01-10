using System;
using System.Linq;
using System.Text.RegularExpressions;
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
                project.HelpFile = project.GetHashCode().ToString();
            }

            return project.HelpFile;
        }

        public static string GetProjectId(IReference reference)
        {
            var projectName = reference.Name;
            var path = reference.FullPath;
            return new QualifiedModuleName(projectName, path, projectName).ProjectId;
        }

        public QualifiedModuleName(IVBProject project)
        {
            _component = null;
            _componentName = null;
            _project = project;
            _projectName = project.Name;
            _projectPath = string.Empty;
            _projectId = GetProjectId(project);
            _projectDisplayName = string.Empty;
            _contentHashCode = 0;
        }

        public QualifiedModuleName(IVBComponent component)
        {
            _project = null; // field is only assigned when the instance refers to a VBProject.

            _component = component;
            _componentName = component.IsWrappingNullReference ? string.Empty : component.Name;

            var components = component.Collection;
            {
                _project = components.Parent;
                _projectName = _project == null ? string.Empty : _project.Name;
                _projectPath = string.Empty;
                _projectId = GetProjectId(_project);
                _projectDisplayName = string.Empty;
            }

            _projectName = _project == null ? string.Empty : _project.Name;
            _projectPath = string.Empty;
            _projectId = GetProjectId(_project);
            _projectDisplayName = string.Empty;

            _contentHashCode = 0;
            if (component.IsWrappingNullReference)
            {
                return;
            }

            var module = component.CodeModule;
            {
                _contentHashCode = module.CountOfLines > 0
                    ? module.GetLines(1, module.CountOfLines).GetHashCode()
                    : 0;
            }
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
            _contentHashCode = 0;
        }

        public QualifiedMemberName QualifyMemberName(string member)
        {
            return new QualifiedMemberName(this, member);
        }

        private readonly IVBComponent _component;
        public IVBComponent Component { get { return _component; } }

        private readonly IVBProject _project;
        public IVBProject Project { get { return _project; } }

        private readonly int _contentHashCode;
        public int ContentHashCode { get { return _contentHashCode; } }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly string _componentName;
        public string ComponentName { get { return _componentName ?? string.Empty; } }

        public string Name { get { return ToString(); } }
        
        private readonly string _projectName;
        public string ProjectName { get { return _projectName; } }

        private readonly string _projectPath;
        public string ProjectPath { get { return _projectPath; } }

        private static readonly Regex CaptionProjectRegex = new Regex(@"^(?:[^-]+)(?:\s-\s)(?<project>.+)(?:\s-\s.*)?$");
        private static readonly Regex OpenModuleRegex = new Regex(@"^(?<project>.+)(?<module>\s-\s\[.*\((Code|UserForm)\)\])$");

        // because this causes a flicker in the VBE, we only want to do it once.
        // we also want to defer it as long as possible because it is only
        // needed in a couple places, and QualifiedModuleName is used in many places.
        private string _projectDisplayName;
        public string ProjectDisplayName
        {
            get
            {
                if (_projectDisplayName != string.Empty)
                {
                    return _projectDisplayName;
                }

                var vbe = _project.VBE;
                var activeProject = vbe.ActiveVBProject;
                var mainWindow = vbe.MainWindow;
                {
                    try
                    {
                        if (_project.HelpFile != activeProject.HelpFile)
                        {
                            vbe.ActiveVBProject = _project;
                        }

                        var caption = mainWindow.Caption;
                        if (CaptionProjectRegex.IsMatch(caption))
                        {
                            caption = CaptionProjectRegex.Matches(caption)[0].Groups["project"].Value;
                            _projectDisplayName = OpenModuleRegex.IsMatch(caption)
                                ? OpenModuleRegex.Matches(caption)[0].Groups["project"].Value
                                : caption;
                        }
                    }
                    // ReSharper disable once EmptyGeneralCatchClause
                    catch { }
                    return _projectDisplayName;
                }
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
