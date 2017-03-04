using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.SourceControl
{
    public abstract class SourceControlProviderBase : ISourceControlProvider
    {
        protected readonly IVBProject Project;

        protected SourceControlProviderBase(IVBProject project)
        {
            Project = project;
        }

        protected SourceControlProviderBase(IVBProject project, IRepository repository)
            :this(project)
        {
            CurrentRepository = repository;
        }

        public IRepository CurrentRepository { get; private set; }
        public abstract IBranch CurrentBranch { get; }
        public abstract IEnumerable<IBranch> Branches { get; }
        public abstract IList<ICommit> UnsyncedLocalCommits { get; }
        public abstract IList<ICommit> UnsyncedRemoteCommits { get; }
        public bool NotifyExternalFileChanges { get; set; }
        public bool HandleVbeSinkEvents { get; set; }
        public abstract IRepository Clone(string remotePathOrUrl, string workingDirectory, SecureCredentials credentials = null);
        public abstract void Push();
        public abstract void Fetch(string remoteName);
        public abstract void AddFile(string filePath);
        public abstract void RemoveFile(string filePath, bool removeFromWorkingDirectory);
        public abstract void CreateBranch(string branch);
        public abstract void DeleteBranch(string branch);
        public abstract IRepository Init(string directory, bool bare = false);
        public abstract void AddOrigin(string path, string trackingBranchName);
        public abstract void Commit(string message);
        public abstract void Publish(string branch);
        public abstract void Unpublish(string branch);
        public abstract bool HasCredentials();
        public abstract IEnumerable<IFileStatusEntry> LastKnownStatus();
        public abstract bool RepoHasRemoteOrigin();

        public virtual void CreateBranch(string sourceBranch, string branch)
        {
            Refresh();
        }

        public virtual IRepository InitVBAProject(string directory)
        {
            if (Project == null)
            {
                throw new SourceControlException(SourceControlText.GitNotInit,
                    new Exception(SourceControlText.NoProjectOpen));
            }

            var projectName = GetProjectNameFromDirectory(directory);
            if (projectName != string.Empty && projectName != Project.Name)
            {
                directory = Path.Combine(directory, Project.Name);
            }

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            NotifyExternalFileChanges = false;
            Project.ExportSourceFiles(directory);
            NotifyExternalFileChanges = true;

            CurrentRepository = new Repository(Project.HelpFile, directory, directory);
            return CurrentRepository;
        }

        public virtual event EventHandler<EventArgs> BranchChanged;

        public virtual void Pull()
        {
            Refresh();
        }

        public virtual void Stage(string filePath)
        {
            NotifyExternalFileChanges = false;
            Project.ExportSourceFiles(CurrentRepository.LocalLocation);
            NotifyExternalFileChanges = true;
        }

        public virtual void Stage(IEnumerable<string> filePaths)
        {
            NotifyExternalFileChanges = false;
            Project.ExportSourceFiles(CurrentRepository.LocalLocation);
            NotifyExternalFileChanges = true;
        }

        public virtual void Merge(string sourceBranch, string destinationBranch)
        {
            Refresh();
        }

        public virtual void Checkout(string branch)
        {
            Refresh();

            var handler = BranchChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public virtual void Undo(string filePath)
        {
            if (File.Exists(filePath))
            {
                var component = Project.VBComponents.OfType<VBComponent>().FirstOrDefault(f => f.Name == filePath.Split('.')[0]);

                HandleVbeSinkEvents = false;
                Project.VBComponents.RemoveSafely(component);
                Project.VBComponents.ImportSourceFile(filePath);
                HandleVbeSinkEvents = true;
            }
        }

        public virtual void Revert()
        {
            Refresh();
        }

        public virtual IEnumerable<IFileStatusEntry> Status()
        {
            NotifyExternalFileChanges = false;
            Project.ExportSourceFiles(CurrentRepository.LocalLocation);
            Project.ImportDocumentTypeSourceFiles(CurrentRepository.LocalLocation);
            NotifyExternalFileChanges = true;
            return null;
        }

        protected string GetProjectNameFromDirectory(string directory)
        {
            var separators = new[] { '/', '\\', '.' };
            return directory.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                            .LastOrDefault(c => c != "git");
        }

        public void ReloadComponent(string filePath)
        {
            HandleVbeSinkEvents = false;
            var moduleName = Path.GetFileNameWithoutExtension(filePath);
            if (moduleName == null)
            {
                return;
            }
            var vbe = Project.VBE;
            var components = Project.VBComponents;
            var pane = vbe.ActiveCodePane;
            {
                var item = components.SingleOrDefault(f => moduleName.Equals(f.Name, StringComparison.InvariantCultureIgnoreCase));
                if (item == null)
                {
                    return;
                }
                if (!pane.IsWrappingNullReference)
                {
                    var module = pane.CodeModule;
                    var component = module.Parent;
                    {
                        var selection = new QualifiedSelection(new QualifiedModuleName(component), pane.Selection);
                        var name = string.IsNullOrEmpty(selection.QualifiedName.ComponentName) ? null : selection.QualifiedName.ComponentName;

                        components.RemoveSafely(item);

                        var directory = CurrentRepository.LocalLocation;
                        directory += directory.EndsWith("\\") ? string.Empty : "\\";
                        components.Import(directory + filePath);

                        VBE.SetSelection(component.Collection.Parent, selection.Selection, name);
                    }
                }
                else
                {
                    components.RemoveSafely(item);

                    var directory = CurrentRepository.LocalLocation;
                    directory += directory.EndsWith("\\") ? string.Empty : "\\";
                    components.Import(directory + filePath);
                }
            }

            HandleVbeSinkEvents = true;
        }

        private void Refresh()
        {
            //Because refreshing removes all components, we need to store the current selection,
            // so we can correctly reset it once the files are imported from the repository.

            HandleVbeSinkEvents = false;

            var vbe = Project.VBE;
            var pane = vbe.ActiveCodePane;
            {
                if (!pane.IsWrappingNullReference)
                {
                    var module = pane.CodeModule;
                    var component = module.Parent;
                    {
                        var selection = new QualifiedSelection(new QualifiedModuleName(component), pane.Selection);
                        var name = string.IsNullOrEmpty(selection.QualifiedName.ComponentName) ? null : selection.QualifiedName.ComponentName;
                        try
                        {
                            Project.LoadAllComponents(CurrentRepository.LocalLocation);
                        }
                        catch (AggregateException ex)
                        {
                            HandleVbeSinkEvents = true;
                            throw new SourceControlException("Unknown exception.", ex);
                        }

                        VBE.SetSelection(component.Collection.Parent, selection.Selection, name);
                    }
                }
                else
                {
                    try
                    {
                        Project.LoadAllComponents(CurrentRepository.LocalLocation);
                    }
                    catch (AggregateException ex)
                    {
                        HandleVbeSinkEvents = true;
                        throw new SourceControlException("Unknown exception.", ex);
                    }
                }
            }

            HandleVbeSinkEvents = true;
        }
    }
}
