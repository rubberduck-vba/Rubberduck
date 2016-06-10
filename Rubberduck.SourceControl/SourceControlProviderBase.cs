using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.SourceControl
{
    public abstract class SourceControlProviderBase : ISourceControlProvider
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        protected readonly VBProject Project;

        protected SourceControlProviderBase(VBProject project)
        {
            Project = project;
        }

        protected SourceControlProviderBase(VBProject project, IRepository repository, ICodePaneWrapperFactory wrapperFactory)
            :this(project)
        {
            CurrentRepository = repository;
            _wrapperFactory = wrapperFactory;
        }

        public IRepository CurrentRepository { get; private set; }
        public abstract IBranch CurrentBranch { get; }
        public abstract IEnumerable<IBranch> Branches { get; }
        public abstract IList<ICommit> UnsyncedLocalCommits { get; }
        public abstract IList<ICommit> UnsyncedRemoteCommits { get; }
        public bool NotifyExternalFileChanges { get; protected set; }
        public bool HandleVbeSinkEvents { get; protected set; }
        public abstract IRepository Clone(string remotePathOrUrl, string workingDirectory);
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

        public virtual void CreateBranch(string sourceBranch, string branch)
        {
            Refresh();
        }

        public virtual IRepository InitVBAProject(string directory)
        {
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
            var componentName = Path.GetFileNameWithoutExtension(filePath);

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

            var codePane = Project.VBE.ActiveCodePane;

            if (codePane != null)
            {
                var codePaneWrapper = _wrapperFactory.Create(codePane);
                var selection = new QualifiedSelection(new QualifiedModuleName(codePaneWrapper.CodeModule.Parent),
                    codePaneWrapper.Selection);
                string name = null;
                if (selection.QualifiedName.Component != null)
                {
                    name = selection.QualifiedName.Component.Name;
                }

                var component = Project.VBComponents.OfType<VBComponent>().FirstOrDefault(f => f.Name == filePath.Split('.')[0]);
                Project.VBComponents.RemoveSafely(component);

                var directory = CurrentRepository.LocalLocation;
                directory += directory.EndsWith("\\") ? string.Empty : "\\";
                Project.VBComponents.Import(directory + filePath);

                Project.VBE.SetSelection(selection.QualifiedName.Project, selection.Selection, name, _wrapperFactory);
            }
            else
            {
                var component = Project.VBComponents.OfType<VBComponent>().FirstOrDefault(f => f.Name == filePath.Split('.')[0]);
                Project.VBComponents.RemoveSafely(component);

                var directory = CurrentRepository.LocalLocation;
                directory += directory.EndsWith("\\") ? string.Empty : "\\";
                Project.VBComponents.Import(directory + filePath);
            }

            HandleVbeSinkEvents = true;
        }

        private void Refresh()
        {
            //Because refreshing removes all components, we need to store the current selection,
            // so we can correctly reset it once the files are imported from the repository.

            HandleVbeSinkEvents = false;

            var codePane = Project.VBE.ActiveCodePane;

            if (codePane != null)
            {
                var codePaneWrapper = _wrapperFactory.Create(codePane);
                var selection = new QualifiedSelection(new QualifiedModuleName(codePaneWrapper.CodeModule.Parent),
                    codePaneWrapper.Selection);
                string name = null;
                if (selection.QualifiedName.Component != null)
                {
                    name = selection.QualifiedName.Component.Name;
                }

                Project.RemoveAllComponents();
                Project.ImportSourceFiles(CurrentRepository.LocalLocation);

                Project.VBE.SetSelection(selection.QualifiedName.Project, selection.Selection, name, _wrapperFactory);
            }
            else
            {
                Project.RemoveAllComponents();
                Project.ImportSourceFiles(CurrentRepository.LocalLocation);
            }

            HandleVbeSinkEvents = true;
        }
    }
}
