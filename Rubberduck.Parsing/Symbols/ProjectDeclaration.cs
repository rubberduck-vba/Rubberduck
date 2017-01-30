using System.Text.RegularExpressions;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectDeclaration : Declaration
    {
        private readonly List<ProjectReference> _projectReferences;

        public ProjectDeclaration(
            QualifiedMemberName qualifiedName,
            string name,
            bool isBuiltIn,
            IVBProject project)
            : base(
                  qualifiedName,
                  null,
                  (Declaration)null,
                  name,
                  null,
                  false,
                  false,
                  Accessibility.Implicit,
                  DeclarationType.Project,
                  null,
                  Selection.Home,
                  false,
                  null,
                  isBuiltIn)
        {
            _project = project;
            _projectReferences = new List<ProjectReference>();
        }

        public ProjectDeclaration(ComProject project, QualifiedModuleName module)
            : this(module.QualifyMemberName(project.Name), project.Name, true, null)
        {
            MajorVersion = project.MajorVersion;
            MinorVersion = project.MinorVersion;
        }

        public long MajorVersion { get; set; }
        public long MinorVersion { get; set; }

        public IReadOnlyList<ProjectReference> ProjectReferences
        {
            get
            {
                return _projectReferences.OrderBy(reference => reference.Priority).ToList();
            }
        }

        private readonly IVBProject _project;
        /// <summary>
        /// Gets a reference to the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public override IVBProject Project { get { return _project; } }

        public void AddProjectReference(string referencedProjectId, int priority)
        {
            if (_projectReferences.Any(p => p.ReferencedProjectId == referencedProjectId))
            {
                return;
            }
            _projectReferences.Add(new ProjectReference(referencedProjectId, priority));
        }

        private static readonly Regex CaptionProjectRegex = new Regex(@"^(?:[^-]+)(?:\s-\s)(?<project>.+)(?:\s-\s.*)?$");
        private static readonly Regex OpenModuleRegex = new Regex(@"^(?<project>.+)(?<module>\s-\s\[.*\((Code|UserForm)\)\])$");

        private string _displayName;
        /// <summary>
        /// WARNING: This property has side effects. It changes the ActiveVBProject, which causes a flicker in the VBE.
        /// This should only be called if it is *absolutely* necessary.
        /// </summary>
        public override string ProjectDisplayName
        {
            get
            {
                if (_displayName != null)
                {
                    return _displayName;
                }

                if (_project == null)
                {
                    _displayName = string.Empty;
                    return _displayName;
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
                            _displayName = OpenModuleRegex.IsMatch(caption)
                                ? OpenModuleRegex.Matches(caption)[0].Groups["project"].Value
                                : caption;
                        }
                    }
                    catch
                    {
                        _displayName = string.Empty;
                    }
                    return _displayName;
                }
            }
        }
    }
}
