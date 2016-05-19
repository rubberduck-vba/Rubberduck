using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;

namespace Rubberduck.Refactorings.ExtractMethod
{
    /// <summary>
    /// A refactoring that extracts a method (procedure or function) 
    /// out of a selection in the active code pane and 
    /// replaces the selected code with a call to the extracted method.
    /// </summary>
    public class ExtractMethodRefactoring : IRefactoring
    {
        private readonly ICodeModuleWrapper _codeModule;
        private Func<QualifiedSelection?, string, IExtractMethodModel> _createMethodModel;
        private IExtractMethodExtraction _extraction;

        public ExtractMethodRefactoring(
            ICodeModuleWrapper codeModule,
            Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel,
            IExtractMethodExtraction extraction)
        {
            _codeModule = codeModule;
            _createMethodModel = createMethodModel;
            _extraction = extraction;

        }

        public void Refactor()
        {
            // TODO : move all this presenter code out
            /*
            var presenter = _factory.Create();
            if (presenter == null)
            {
                OnInvalidSelection();
                return;
            }

            */
            var qualifiedSelection = _codeModule.QualifiedSelection;
            if (qualifiedSelection == null)
            {
                return;
            }

            var selection = qualifiedSelection.Value.Selection;
            var selectedCode = _codeModule.GetLines(selection);
            var model = _createMethodModel(qualifiedSelection, selectedCode);
            if (model == null)
            {
                return;
            }

            /*
            var success = presenter.Show(model,_createProc);
            if (!success)
            {
                return;
            }
            */

            _extraction.apply(_codeModule, model, selection);
        }

        public void Refactor(QualifiedSelection target)
        {
            _codeModule.SetSelection(target);
            Refactor();
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