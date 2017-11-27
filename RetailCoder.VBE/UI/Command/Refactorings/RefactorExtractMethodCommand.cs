using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Refactorings;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRefactoringFactory<ExtractMethodRefactoring> _refactoringFactory;

        public RefactorExtractMethodCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            IRefactoringFactory<ExtractMethodRefactoring> refactoringFactory)
            : base (vbe)
        {
            _state = state;
            _refactoringFactory = refactoringFactory;
        }

        public override RubberduckHotkey Hotkey => RubberduckHotkey.RefactorExtractMethod;

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var pane = Vbe.ActiveCodePane;
            var module = pane.CodeModule;
            {
                var qualifiedSelection = pane.GetQualifiedSelection();
                if (!qualifiedSelection.HasValue || module.IsWrappingNullReference)
                {
                    return false;
                }
                
                var validator = new ExtractMethodSelectionValidation(_state.AllDeclarations);
                var canExecute = validator.ValidateSelection(qualifiedSelection.Value);

                return canExecute;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var qualifiedSelection = Vbe.ActiveCodePane.GetQualifiedSelection();

            if (qualifiedSelection == null)
            {
                return;
            }

            var pane = Vbe.ActiveCodePane;
            if (pane == null)
            {
                return;
            }

            var validator = new ExtractMethodSelectionValidation(_state.AllDeclarations);
            var canExecute = validator.ValidateSelection(qualifiedSelection.Value);

            if (!canExecute)
            {
                return;
            }

            var module = pane.CodeModule;
            var selection = module.GetQualifiedSelection();

            if (selection == null)
            {
                return;
            }

            /* TODO: Refactor the section to make command ignorant of data
             * This section needs to be refactored. The way it is, the command knows too much
             * about the validator and the refactoring. Getting data from validator should
             * be refactoring's responsibility, which implies the validation is refactoring's
             * responsiblity. Note where indicated.
             */
            

            var refactoring = _refactoringFactory.Create();
            refactoring.Validator = validator; //TODO: Refactor
            refactoring.Refactor(selection.Value);
            _refactoringFactory.Release(refactoring);

            /*
            using (var view = new ExtractMethodDialog(new ExtractMethodViewModel()))
            {
                var factory = new ExtractMethodPresenterFactory(Vbe, view, _indenter, _state, qualifiedSelection.Value);
                var refactoring = new ExtractMethodRefactoring(Vbe, module, factory);
                refactoring.Refactor(qualifiedSelection.Value);
            }
            */

            /*
            {
                Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = (qs, code) =>
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
                };

                var extraction = new ExtractMethodExtraction();
                // bug: access to disposed closure - todo: make ExtractMethodRefactoring request reparse like everyone else.
                Action<object> parseRequest = obj => _state.OnParseRequested(obj); 

                var refactoring = new ExtractMethodRefactoring(module, parseRequest, createMethodModel, extraction);
                refactoring.InvalidSelection += HandleInvalidSelection;
                refactoring.Refactor();
            }
            */
        }
    }
}
