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
        protected VBProject Project;

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
        public abstract IRepository Clone(string remotePathOrUrl, string workingDirectory);
        public abstract void Push();
        public abstract void Fetch(string remoteName);
        public abstract void AddFile(string filePath);
        public abstract void RemoveFile(string filePath);
        public abstract void CreateBranch(string branch);
        public abstract void DeleteBranch(string branch);
        public abstract IRepository Init(string directory, bool bare = false);
        public abstract void Commit(string message);
        public abstract void Publish(string branch);
        public abstract void Unpublish(string branch);

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

            Project.ExportSourceFiles(directory);
            CurrentRepository = new Repository(Project.Name, directory, directory);
            return CurrentRepository;
        }

        public virtual event EventHandler<EventArgs> BranchChanged;

        public virtual void Pull()
        {
            Refresh();
        }

        public virtual void Stage(string filePath)
        {
            Project.ExportSourceFiles(CurrentRepository.LocalLocation);
        }

        public virtual void Stage(IEnumerable<string> filePaths)
        {
            Project.ExportSourceFiles(CurrentRepository.LocalLocation);
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
                var component = Project.VBComponents.Item(componentName);
                Project.VBComponents.RemoveSafely(component);
                Project.VBComponents.ImportSourceFile(filePath);
            }
        }

        public virtual void Revert()
        {
            Refresh();
        }

        public virtual IEnumerable<IFileStatusEntry> Status()
        {
            Project.ExportSourceFiles(CurrentRepository.LocalLocation);
            return null;
        }

        protected string GetProjectNameFromDirectory(string directory)
        {
            var separators = new[] { '/', '\\', '.' };
            return directory.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                            .LastOrDefault(c => c != "git");
        }

        private void Refresh()
        {
            //Because refreshing removes all components, we need to store the current selection,
            // so we can correctly reset it once the files are imported from the repository.

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
        }
    }
}
