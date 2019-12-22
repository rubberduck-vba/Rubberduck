using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Parsing.Common;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [Disabled]
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private readonly IVBE _vbe;

        public RefactorExtractMethodCommand(IVBE vbe, RubberduckParserState state, IIndenter indenter)
        {
            _state = state;
            _indenter = indenter;
            _vbe = vbe;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            var qualifiedSelection = _vbe.GetActiveSelection();
            if (!qualifiedSelection.HasValue)
            {
                return false;
            }
            if (_state.IsNewOrModified(qualifiedSelection.Value.QualifiedName))
            {
                return false;
            }

            var allDeclarations = _state.AllDeclarations;
            var extractMethodValidation = new ExtractMethodSelectionValidation(allDeclarations);
                
            var canExecute = extractMethodValidation.withinSingleProcedure(qualifiedSelection.Value);

            return canExecute;
        }

        protected override void OnExecute(object parameter)
        {
            var declarations = _state.AllDeclarations;
            var qualifiedSelection = _vbe.GetActiveSelection();

            var extractMethodValidation = new ExtractMethodSelectionValidation(declarations);
            var canExecute = extractMethodValidation.withinSingleProcedure(qualifiedSelection.Value);
            if (!canExecute)
            {
                return;
            }

            using (var pane = _vbe.ActiveCodePane)
            using (var module = pane.CodeModule)
            {
                var extraction = new ExtractMethodExtraction();
                // bug: access to disposed closure

                // todo: make ExtractMethodRefactoring request reparse like everyone else.
                var refactoring = new ExtractMethodRefactoring(module, ParseRequest, CreateMethodModel, extraction);
                refactoring.InvalidSelection += HandleInvalidSelection;
                refactoring.Refactor();


                void ParseRequest(object obj) => _state.OnParseRequested(obj);

                IExtractMethodModel CreateMethodModel(QualifiedSelection? qs, string code)
                {
                    if (qs == null)
                    {
                        return null;
                    }
                    //TODO: Pull these even further back;
                    //      and implement with IProvider<IExtractMethodRule>
                    var rules = new List<IExtractMethodRule>
                    {
                        new ExtractMethodRuleInSelection(),
                        new ExtractMethodRuleIsAssignedInSelection(),
                        new ExtractMethodRuleUsedAfter(),
                        new ExtractMethodRuleUsedBefore()
                    };

                    var paramClassify = new ExtractMethodParameterClassification(rules);

                    var extractedMethod = new ExtractedMethod();
                    var extractedMethodModel = new ExtractMethodModel(extractedMethod, paramClassify);
                    extractedMethodModel.extract(declarations, qs.Value, code);
                    return extractedMethodModel;
                }
            }
        }

        private void HandleInvalidSelection(object sender, EventArgs e)
        {
            MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}
