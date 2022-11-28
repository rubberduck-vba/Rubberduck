using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.ExtractMethod
{
    /// <summary>
    /// A refactoring that extracts a method (procedure or function) 
    /// out of a selection in the active code pane and 
    /// replaces the selected code with a call to the extracted method.
    /// </summary>
    public class ExtractMethodRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly ICodeModule _codeModule;
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;

        // TODO: encapsulate the model. See the TODO in command class
        private QualifiedSelection selection;
        public ExtractMethodSelectionValidation Validator;

        public ExtractMethodRefactoring(
            IVBE vbe,
            IIndenter indenter,
            RubberduckParserState state
        )
        {
            _vbe = vbe;
            _codeModule = _vbe.ActiveCodePane.CodeModule;
            _indenter = indenter;
            _state = state;
        }

        public void Refactor()
        {
            if (!_codeModule.GetQualifiedSelection().HasValue)
            {
                OnInvalidSelection();
                return;
            }

            selection = _codeModule.GetQualifiedSelection().Value;
            
            var model = new ExtractMethodModel(_state, selection, Validator.SelectedContexts, _indenter, _codeModule);
            var presenter = ExtractMethodPresenter.Create(model);

            if (presenter == null)
            {
                return;
            }

            model = presenter.Show();
            if (model == null)
            {
                return;
            }

            QualifiedSelection? oldSelection;
            if (!_codeModule.IsWrappingNullReference)
            {
                oldSelection = _codeModule.GetQualifiedSelection();
            }
            else
            {
                return;
            }

            if (oldSelection.HasValue)
            {
                _codeModule.CodePane.Selection = oldSelection.Value.Selection;
            }

            model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _codeModule.CodePane;
            {
                pane.Selection = target.Selection;
                Refactor();
            }
        }

        public void Refactor(Declaration target)
        {
            OnInvalidSelection();
        }

        private void ExtractMethod()
        {

            #region to be put back when allow subs and functions
            /* Remove this entirely for now.
            // assumes these are declared *before* the selection...
            var offset = 0;
            foreach (var declaration in model.DeclarationsToMove.OrderBy(e => e.Selection.StartLine))
            {
                var target = new Selection(
                    declaration.Selection.StartLine - offset,
                    declaration.Selection.StartColumn,
                    declaration.Selection.EndLine - offset,
                    declaration.Selection.EndColumn);

                _codeModule.DeleteLines(target);
                offset += declaration.Selection.LineCount;
            }
            */
            #endregion

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

    }
}
