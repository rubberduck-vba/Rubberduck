using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.VBEditor.ComManagement.NonDisposalDecorators
{
    public class VbeNonDisposalDecorator<T> : NonDisposalDecoratorBase<T>, IVBE
        where T: IVBE
    {
        public VbeNonDisposalDecorator(T vbe)
            : base(vbe)
        {}

        public bool Equals(IVBE other)
        {
            return WrappedItem.Equals(other);
        }

        public VBEKind Kind => WrappedItem.Kind;
        public string Version => WrappedItem.Version;
        public object HardReference => WrappedItem.HardReference;
        public IWindow ActiveWindow => WrappedItem.ActiveWindow;

        public ICodePane ActiveCodePane
        {
            get => WrappedItem.ActiveCodePane;
            set => WrappedItem.ActiveCodePane = value;
        }

        public IVBProject ActiveVBProject
        {
            get => WrappedItem.ActiveVBProject;
            set => WrappedItem.ActiveVBProject = value;
        }

        public IVBComponent SelectedVBComponent => WrappedItem.SelectedVBComponent;
        public IWindow MainWindow => WrappedItem.MainWindow;
        public IAddIns AddIns => WrappedItem.AddIns;
        public IVBProjects VBProjects => WrappedItem.VBProjects;
        public ICodePanes CodePanes => WrappedItem.CodePanes;
        public ICommandBars CommandBars => WrappedItem.CommandBars;
        public IWindows Windows => WrappedItem.Windows;

        public IHostApplication HostApplication()
        {
            return WrappedItem.HostApplication();
        }

        public IWindow ActiveMDIChild()
        {
            return WrappedItem.ActiveMDIChild();
        }

        public QualifiedSelection? GetActiveSelection()
        {
            return WrappedItem.GetActiveSelection();
        }

        public bool IsInDesignMode => WrappedItem.IsInDesignMode;

        public int ProjectsCount => WrappedItem.ProjectsCount;

        public ITempSourceFileHandler TempSourceFileHandler => WrappedItem.TempSourceFileHandler;
    }
}