using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;

namespace Rubberduck.UI
{
    /// <summary>
    /// ProjectToExportFolderMap is a singleton container of Project/Folder pairs to
    /// support multiple instances of the ExportAllCommand class
    /// </summary>
    public class ProjectToExportFolderMap
    {
        private readonly Dictionary<string, string> _projectToExportFolderMap;

        public ProjectToExportFolderMap()
        {
            _projectToExportFolderMap = new Dictionary<string, string>();
        }

        public void AssignProjectExportFolder(IVBProject project, string exportFolderpath)
        {
            if (project is null || string.IsNullOrWhiteSpace(exportFolderpath))
            {
                return;
            }

            if (!_projectToExportFolderMap.ContainsKey(project.FileName))
            {
                _projectToExportFolderMap.Add(project.FileName, exportFolderpath);
                return;
            }

            _projectToExportFolderMap[project.FileName] = exportFolderpath;
        }

        public bool TryGetExportPathForProject(IVBProject project, out string exportFolderpath)
        {
            exportFolderpath = string.Empty;
            if (string.IsNullOrWhiteSpace(project.FileName))
            {
                return false;
            }

            return _projectToExportFolderMap.TryGetValue(project.FileName, out exportFolderpath);
        }

        public void RemoveProject(IVBProject project)
        {
            _projectToExportFolderMap.Remove(project.FileName);
        }
    }
}
