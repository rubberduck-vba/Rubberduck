using System;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactoring
{
    /// <summary>
    /// A refactoring that extracts a method (procedure or function) 
    /// out of a selection in the active code pane and 
    /// replaces the selected code with a call to the extracted method.
    /// </summary>
    public class ExtractMethodRefactoring : IRefactoring
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly Declarations _declarations;

        public ExtractMethodRefactoring(IActiveCodePaneEditor editor, Declarations declarations)
        {
            _editor = editor;
            _declarations = declarations;
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        /// <summary>
        /// Returns the <see cref="Declaration"/> for the procedure 
        /// that contains the specified <see cref="QualifiedSelection"/>.
        /// </summary>
        public Declaration AcquireTarget(QualifiedSelection selection)
        {
            var scope = _editor.GetSelectedProcedureScope(selection.Selection);
            if (string.IsNullOrEmpty(scope))
            {
                return null;
            }

            return _declarations.Items.SingleOrDefault(declaration => 
                declaration.Project == selection.QualifiedName.Project
                && declaration.Scope == scope 
                && ProcedureTypes.Contains(declaration.DeclarationType));
        }

        public void Refactor()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                OnInvalidSelection();
                return;
            }
            
            Refactor(selection.Value);
        }

        /// <summary>
        /// An event that is raised when refactoring is not possible due to an invalid selection.
        /// </summary>
        public event EventHandler InvalidSelection;

        private void OnInvalidSelection()
        {
            var handler = InvalidSelection;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public void Refactor(QualifiedSelection selection)
        {
            var member = AcquireTarget(selection);
            if (member == null)
            {
                OnInvalidSelection();
                return;
            }

            using (var view = new ExtractMethodDialog())
            {
                var presenter = new ExtractMethodPresenter(_editor, view, member, selection, _declarations);
                presenter.Show();
            }
        }

        void IRefactoring.Refactor(Declaration target)
        {
            throw new NotImplementedException("This refactoring requires a QualifiedSelection.");
        }
    }
}