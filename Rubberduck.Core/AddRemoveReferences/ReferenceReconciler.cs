using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Interaction;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.AddRemoveReferences
{
    public interface IReferenceReconciler
    {
        void ReconcileReferences(IAddRemoveReferencesModel model);
        List<ReferenceModel> ReconcileReferences(IAddRemoveReferencesModel model, List<ReferenceModel> allReferences);
        ReferenceModel TryAddReference(IVBProject project, string path);
        ReferenceModel TryAddReference(IVBProject project, ReferenceModel reference);
        ReferenceModel GetLibraryInfoFromPath(string path);
    }

    public class ReferenceReconciler : IReferenceReconciler
    {
        private readonly IMessageBox _messageBox;
        private readonly IConfigProvider<GeneralSettings> _settings;
        private readonly IComLibraryProvider _tlbProvider;

        public ReferenceReconciler(IMessageBox messageBox, IConfigProvider<GeneralSettings> settings, IComLibraryProvider tlbProvider)
        {
            _messageBox = messageBox;
            _settings = settings;
            _tlbProvider = tlbProvider;
        }

        public void ReconcileReferences(IAddRemoveReferencesModel model)
        {
            ReconcileReferences(model, model.NewReferences.ToList());
        }

        public List<ReferenceModel> ReconcileReferences(IAddRemoveReferencesModel model, List<ReferenceModel> allReferences)
        {
            var selected = allReferences.Where(reference => !reference.IsBuiltIn && reference.Priority.HasValue)
                .ToDictionary(reference => reference.FullPath);

            var output = selected.Values.Where(reference => reference.IsBuiltIn).ToList();

            var project = model.Project.Project;
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

            return output;
        }

        public ReferenceModel GetLibraryInfoFromPath(string path)
        {
            try
            {
                return new ReferenceModel(_tlbProvider.LoadTypeLibrary(path));
            }
            catch
            {
                // Most likely this is a project. If not, it we can't fail here because it could have come from the Apply
                // button in the AddRemoveReferencesDialog. Wait for it...  :-P
                return new ReferenceModel(path);
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
                        return reference is null ? null : new ReferenceModel(reference, references.Count);
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
    }
}
