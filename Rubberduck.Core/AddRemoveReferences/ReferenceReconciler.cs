using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AddRemoveReferences
{
    public interface IReferenceReconciler
    {
        List<ReferenceModel> ReconcileReferences(IAddRemoveReferencesModel model);
        List<ReferenceModel> ReconcileReferences(IAddRemoveReferencesModel model, List<ReferenceModel> allReferences);
        ReferenceModel TryAddReference(IVBProject project, string path);
        ReferenceModel TryAddReference(IVBProject project, ReferenceModel reference);
        ReferenceModel GetLibraryInfoFromPath(string path);
        void UpdateSettings(IAddRemoveReferencesModel model, bool recent = false);
    }

    public class ReferenceReconciler : IReferenceReconciler
    {
        public static readonly List<string> TypeLibraryExtensions = new List<string> { ".olb", ".tlb", ".dll", ".ocx", ".exe" };

        private readonly IMessageBox _messageBox;
        private readonly IConfigurationService<ReferenceSettings> _settings;
        private readonly IComLibraryProvider _libraryProvider;
        private readonly IProjectsProvider _projectsProvider;

        public ReferenceReconciler(
            IMessageBox messageBox, 
            IConfigurationService<ReferenceSettings> settings, 
            IComLibraryProvider libraryProvider,
            IProjectsProvider projectsProvider)
        {
            _messageBox = messageBox;
            _settings = settings;
            _libraryProvider = libraryProvider;
            _projectsProvider = projectsProvider;
        }

        public List<ReferenceModel> ReconcileReferences(IAddRemoveReferencesModel model)
        {
            if (model?.NewReferences is null || !model.NewReferences.Any())
            {
                return new List<ReferenceModel>();
            }

            return ReconcileReferences(model, model.NewReferences.ToList());
        }

        //TODO test for simple adds.
        public List<ReferenceModel> ReconcileReferences(IAddRemoveReferencesModel model, List<ReferenceModel> allReferences)
        {
            if (model is null || allReferences is null || !allReferences.Any())
            {
                return new List<ReferenceModel>();
            }

            var project = _projectsProvider.Project(model.Project.ProjectId);

            if (project == null)
            {
                return new List<ReferenceModel>();
            }

            var selected = allReferences.Where(reference => !reference.IsBuiltIn && reference.Priority.HasValue)
                .ToDictionary(reference => reference.FullPath);

            var output = selected.Values.Where(reference => reference.IsBuiltIn).ToList();

            using (var references = project.References)
            {
                foreach (var reference in references)
                {
                    try
                    {
                        if (!reference.IsBuiltIn)
                        {
                            references.Remove(reference);
                        }
                    }
                    finally
                    {
                        reference.Dispose();                        
                    }
                }
                output.AddRange(selected.Values.OrderBy(selection => selection.Priority)
                    .Select(reference => TryAddReference(project, reference)).Where(added => added != null));
            }

            UpdateSettings(model, true);
            return output;
        }

        public ReferenceModel GetLibraryInfoFromPath(string path)
        {
            try
            {
                var extension = Path.GetExtension(path)?.ToLowerInvariant() ?? string.Empty;
                if (string.IsNullOrEmpty(extension))
                {
                    return null;
                }

                // LoadTypeLibrary will attempt to open files in the host, so only attempt on possible COM servers.
                if (TypeLibraryExtensions.Contains(extension))
                {
                    var type = _libraryProvider.LoadTypeLibrary(path);
                    return new ReferenceModel(path, type, _libraryProvider);
                }
                return new ReferenceModel(path);
            }
            catch
            {
                // Most likely this is unloadable. If not, it we can't fail here because it could have come from the Apply
                // button in the AddRemoveReferencesDialog. Wait for it...  :-P
                return new ReferenceModel(path, true);
            }
        }

        public ReferenceModel TryAddReference(IVBProject project, string path)
        {
            using (var references = project.References)
            {
                try
                {
                    using (var reference = references.AddFromFile(path))
                    {
                        return reference is null ? null : new ReferenceModel(reference, references.Count) { IsRecent = true };
                    }
                }
                catch (COMException ex)
                {
                    _messageBox.NotifyWarn(ex.Message, RubberduckUI.References_AddFailedCaption);
                }
                return null;
            }
        }

        public ReferenceModel TryAddReference(IVBProject project, ReferenceModel reference)
        {
            using (var references = project.References)
            {
                try
                {
                    using (references.AddFromFile(reference.FullPath))
                    {
                        reference.Priority = references.Count;
                        reference.IsRecent = true;
                        return reference;
                    }
                }
                catch (COMException ex)
                {
                    _messageBox.NotifyWarn(ex.Message, RubberduckUI.References_AddFailedCaption);
                }
                return null;
            }
        }

        public void UpdateSettings(IAddRemoveReferencesModel model, bool recent = false)
        {
            if (model?.Settings is null || model.References is null)
            {
                return;
            }

            if (recent)
            {
                model.Settings.UpdateRecentReferencesForHost(model.HostApplication,
                    model.References.Where(reference => reference.IsReferenced && !reference.IsBuiltIn)
                        .Select(reference => reference.ToReferenceInfo()).ToList());
                
            }

            model.Settings.UpdatePinnedReferencesForHost(model.HostApplication,
                model.References.Where(reference => reference.IsPinned).Select(reference => reference.ToReferenceInfo())
                    .ToList());

            _settings.Save(model.Settings);
        }
    }
}
