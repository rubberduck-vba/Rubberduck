using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.AddRemoveReferences
{
    public interface IAddRemoveReferencesPresenterFactory
    {
        AddRemoveReferencesPresenter Create(ProjectDeclaration projectDeclaration);
    }

    public class AddRemoveReferencesPresenterFactory : IAddRemoveReferencesPresenterFactory
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly bool _use64BitPaths = Environment.Is64BitProcess;

        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IConfigurationService<ReferenceSettings> _settings;
        private readonly IRegisteredLibraryFinderService _finder;
        private readonly IReferenceReconciler _reconciler;
        private readonly IFileSystemBrowserFactory _browser;
        private readonly IProjectsProvider _projectsProvider;

        public AddRemoveReferencesPresenterFactory(IVBE vbe,
            RubberduckParserState state,
            IConfigurationService<ReferenceSettings> settingsProvider, 
            IRegisteredLibraryFinderService finder,
            IReferenceReconciler reconciler,
            IFileSystemBrowserFactory browser,
            IProjectsProvider projectsProvider)
        {
            _vbe = vbe;
            _state = state;
            _settings = settingsProvider;
            _finder = finder;
            _reconciler = reconciler;
            _browser = browser;
            _projectsProvider = projectsProvider;
        }

        public AddRemoveReferencesPresenter Create(ProjectDeclaration projectDeclaration)
        {
            if (projectDeclaration is null)
            {
                return null;
            }

            var project = _projectsProvider.Project(projectDeclaration.ProjectId);

            if (project == null)
            {
                return null;
            }

            AddRemoveReferencesModel model = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                var refs = new Dictionary<RegisteredLibraryKey, RegisteredLibraryInfo>();
                // Iterating the returned libraries here instead of just .ToDictionary() using because we can't trust that the registry doesn't contain errors.
                foreach (var reference in _finder.FindRegisteredLibraries())
                {
                    if (refs.ContainsKey(reference.UniqueId))
                    {
                        _logger.Warn(
                            $"Duplicate registry definition for {reference.Guid} version {reference.Version}.");
                        continue;
                    }

                    refs.Add(reference.UniqueId, reference);
                }

                var models = new Dictionary<RegisteredLibraryKey, ReferenceModel>();
                using (var references = project.References)
                {
                    if (references is null)
                    {
                        return null;
                    }

                    var priority = 1;
                    foreach (var reference in references)
                    {
                        var guid = Guid.TryParse(reference.Guid, out var result) ? result : Guid.Empty;

                        // This avoids collisions when the parse actually succeeds, but the result is empty.
                        if (guid.Equals(Guid.Empty))
                        {
                            guid = new Guid(priority, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
                        }

                        var libraryId = new RegisteredLibraryKey(guid, reference.Major, reference.Minor);

                        // TODO: If for some reason the VBA reference is broken, we could technically use this to repair it. Just a thought...
                        var adding = refs.ContainsKey(libraryId)
                            ? new ReferenceModel(refs[libraryId], reference, priority++)
                            : new ReferenceModel(reference, priority++);

                        adding.IsUsed = adding.IsBuiltIn ||
                                        _state.DeclarationFinder.IsReferenceUsedInProject(projectDeclaration,
                                            adding.ToReferenceInfo());

                        models.Add(libraryId, adding);
                        reference.Dispose();
                    }
                }

                foreach (var reference in refs.Where(library =>
                    (_use64BitPaths || library.Value.Has32BitVersion) &&
                    !models.ContainsKey(library.Key)))
                {
                    models.Add(reference.Key, new ReferenceModel(reference.Value));
                }

                var settings = _settings.Read();
                model = new AddRemoveReferencesModel(_state, projectDeclaration, models.Values, settings);
                if (AddRemoveReferencesViewModel.HostHasProjects)
                {
                    model.References.AddRange(GetUserProjectFolderModels(model.Settings).Where(proj =>
                        !model.References.Any(item =>
                            item.FullPath.Equals(proj.FullPath, StringComparison.OrdinalIgnoreCase))));
                }
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "Unexpected exception attempting to create AddRemoveReferencesModel.");
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

            return (model != null)
                ? new AddRemoveReferencesPresenter(
                    new AddRemoveReferencesDialog(new AddRemoveReferencesViewModel(model, _reconciler, _browser, _projectsProvider)))
                : null;                 
        }

        private IEnumerable<ReferenceModel> GetUserProjectFolderModels(IReferenceSettings settings)
        {
            var host = Path.GetFileName(Application.ExecutablePath).ToUpperInvariant();
            if (!AddRemoveReferencesViewModel.HostFileFilters.ContainsKey(host))
            {
                return Enumerable.Empty<ReferenceModel>();
            }

            var filter = AddRemoveReferencesViewModel.HostFileFilters[host].Select(ext => $".{ext}").ToList();
            var projects = new List<ReferenceModel>();
            
            foreach (var path in settings.ProjectPaths.Where(Directory.Exists))
            {
                try
                {
                    foreach (var fullPath in Directory.EnumerateFiles(path))
                    {
                        try
                        {
                            if (!filter.Contains(Path.GetExtension(fullPath)))
                            {
                                continue;                               
                            }
                            var file = Path.GetFileName(fullPath);

                            projects.Add(new ReferenceModel(
                                new ReferenceInfo(Guid.Empty, file, fullPath, 0, 0),
                                ReferenceKind.Project, 
                                settings.IsRecentProject(fullPath, host),
                                settings.IsPinnedProject(fullPath, host)));
                        }
                        catch
                        {
                            // 'Ignored
                        }
                    }
                }
                catch
                {
                    _logger.Info("User project directory in reference settings does not exist.");
                }
            }

            return projects;
        }

        public AddRemoveReferencesPresenter Create()
        {
            var selectedProject = SelectedProjectDeclaration();
            return selectedProject is null
                ? null
                : Create(selectedProject);
        }

        private ProjectDeclaration SelectedProjectDeclaration()
        {
            var projectId = SelectedProjectId();

            if (projectId == null)
            {
                return null;
            }

            return _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .OfType<ProjectDeclaration>()
                .FirstOrDefault(item => item.ProjectId.Equals(projectId));
        }

        private string SelectedProjectId()
        {
            using (var selectedComponent = _vbe.SelectedVBComponent)
            {
                using (var project = selectedComponent.ParentProject)
                {

                    if (project == null || project.IsWrappingNullReference)
                    {
                        return null;
                    }

                    return project.ProjectId;
                }
            }
        }
    }
}
