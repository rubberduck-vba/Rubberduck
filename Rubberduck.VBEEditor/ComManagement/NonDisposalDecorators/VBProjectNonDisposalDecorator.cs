using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.NonDisposalDecorators
{
    public class VBProjectNonDisposalDecorator<T> : NonDisposalDecoratorBase<T>, IVBProject
        where T : IVBProject
    {
        public VBProjectNonDisposalDecorator(T project)
            : base(project)
        { }

        public bool Equals(IVBProject other)
        {
            return WrappedItem.Equals(other);
        }

        public IApplication Application => WrappedItem.Application;

        public IApplication Parent => WrappedItem.Parent;

        public IVBE VBE => WrappedItem.VBE;

        public IVBProjects Collection => WrappedItem.Collection;

        public IReferences References => WrappedItem.References;

        public IVBComponents VBComponents => WrappedItem.VBComponents;

        public string ProjectId => WrappedItem.ProjectId;

        public string Name
        {
            get => WrappedItem.Name;
            set => WrappedItem.Name = value;
        }

        public string Description
        {
            get => WrappedItem.Description;
            set => WrappedItem.Description = value;
        }

        public string HelpFile
        {
            get => WrappedItem.HelpFile;
            set => WrappedItem.HelpFile = value;
        }

        public string FileName => WrappedItem.FileName;

        public string BuildFileName => WrappedItem.BuildFileName;

        public bool IsSaved => WrappedItem.IsSaved;

        public ProjectType Type => WrappedItem.Type;

        public EnvironmentMode Mode => WrappedItem.Mode;

        public ProjectProtection Protection => WrappedItem.Protection;

        public void AssignProjectId()
        {
            WrappedItem.AssignProjectId();
        }

        public void SaveAs(string fileName)
        {
            WrappedItem.SaveAs(fileName);
        }

        public void MakeCompiledFile()
        {
            WrappedItem.MakeCompiledFile();
        }

        public void ExportSourceFiles(string folder)
        {
            WrappedItem.ExportSourceFiles(folder);
        }

        public string ProjectDisplayName => WrappedItem.ProjectDisplayName;

        public IReadOnlyList<string> ComponentNames()
        {
            return WrappedItem.ComponentNames();
        }
    }
}