using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.VBEditor
{
    public struct QualifiedSelection
    {
        public QualifiedSelection(QualifiedModuleName qualifiedName, Selection selection, IRubberduckCodePaneFactory factory)
        {
            _qualifiedName = qualifiedName;
            _selection = selection;
            _factory = factory;
        }

        private readonly IRubberduckCodePaneFactory _factory;

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get {return _qualifiedName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        /// <summary>
        /// Sets the current selection in the VBE.
        /// </summary>
        public void Select()
        {
            var codePane = _factory.Create(_qualifiedName.Component.CodeModule.CodePane);
            codePane.Selection = _selection;
        }

        public override string ToString()
        {
            return string.Concat(QualifiedName, " ", Selection);
        }
    }
}