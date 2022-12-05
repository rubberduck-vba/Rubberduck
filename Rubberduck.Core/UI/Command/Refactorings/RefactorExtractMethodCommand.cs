using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Refactorings;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorExtractMethodCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRefactoringFactory<ExtractMethodRefactoring> _refactoringFactory;
        private readonly IVBE _vbe;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly ISelectionProvider _selectionProvider;

        public RefactorExtractMethodCommand(
            ExtractMethodRefactoring refactoring, 
            ExtractMethodFailedNotifier failureNotifier, 
            RubberduckParserState state, 
            ISelectionProvider selectionProvider, 
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(refactoring, failureNotifier, selectionProvider, state)
        {
            _state = state;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _selectionProvider = selectionProvider;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }
        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var member = _selectedDeclarationProvider.SelectedDeclaration();
            var moduleContext = _selectedDeclarationProvider.SelectedModule().Context;
            var moduleName = _selectedDeclarationProvider.SelectedModule().QualifiedModuleName;

            if (member == null || _state.IsNewOrModified(member.QualifiedModuleName) || !_selectionProvider.Selection(moduleName).HasValue)
            {
                return false;
            }

            return true;
            //var parameters = _state.DeclarationFinder
            //    .UserDeclarations(DeclarationType.Parameter)
            //    .Where(item => member.Equals(item.ParentScopeDeclaration))
            //    .ToList();

            //return member.DeclarationType == DeclarationType.PropertyLet
            //        || member.DeclarationType == DeclarationType.PropertySet
            //    ? parameters.Count > 2
            //    : parameters.Count > 1;
        }

        //protected bool EvaluateCanExecute(object parameter)
        //{
        //    if (_vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
        //    {
        //        return false;
        //    }

        //    var pane = _vbe.ActiveCodePane;
        //    var module = pane.CodeModule;
        //    {
        //        var qualifiedSelection = pane.GetQualifiedSelection();
        //        if (!qualifiedSelection.HasValue || module.IsWrappingNullReference)
        //        {
        //            return false;
        //        }
                
        //        var validator = new ExtractMethodSelectionValidation(_state.AllDeclarations, module);
        //        var canExecute = validator.ValidateSelection(qualifiedSelection.Value);

        //        return canExecute;
        //    }
        //}

        //protected override void OnExecute(object parameter)
        //{
        //    var qualifiedSelection = _vbe.ActiveCodePane.GetQualifiedSelection();

        //    if (qualifiedSelection == null)
        //    {
        //        return;
        //    }

        //    var pane = _vbe.ActiveCodePane;
        //    if (pane == null)
        //    {
        //        return;
        //    }
            
        //    var module = pane.CodeModule;
        //    var selection = module.GetQualifiedSelection();

        //    if (selection == null)
        //    {
        //        return;
        //    }

        //    var validator = new ExtractMethodSelectionValidation(_state.AllDeclarations, module);
        //    var canExecute = validator.ValidateSelection(qualifiedSelection.Value);

        //    if (!canExecute)
        //    {
        //        return;
        //    }

        //    /* TODO: Refactor the section to make command ignorant of data
        //     * This section needs to be refactored. The way it is, the command knows too much
        //     * about the validator and the refactoring. Getting data from validator should
        //     * be refactoring's responsibility, which implies the validation is refactoring's
        //     * responsiblity. Note where indicated.
        //     */


        //    var refactoring = _refactoringFactory.Create();
        //    refactoring.Validator = validator; //TODO: Refactor
        //    refactoring.Refactor(selection.Value);
        //    _refactoringFactory.Release(refactoring);

        //    /*
        //    using (var view = new ExtractMethodDialog(new ExtractMethodViewModel()))
        //    {
        //        var factory = new ExtractMethodPresenterFactory(Vbe, view, _indenter, _state, qualifiedSelection.Value);
        //        var refactoring = new ExtractMethodRefactoring(Vbe, module, factory);
        //        refactoring.Refactor(qualifiedSelection.Value);
        //    }
        //    */

        //    /*
        //    {
        //        Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = (qs, code) =>
        //        {
        //            if (qs == null)
        //            {
        //                return null;
        //            }
        //            //TODO: Pull these even further back;
        //            //      and implement with IProvider<IExtractMethodRule>
        //            var rules = new List<IExtractMethodRule>
        //            {
        //                new ExtractMethodRuleInSelection(),
        //                new ExtractMethodRuleIsAssignedInSelection(),
        //                new ExtractMethodRuleUsedAfter(),
        //                new ExtractMethodRuleUsedBefore()
        //            };

        //            var paramClassify = new ExtractMethodParameterClassification(rules);

        //            var extractedMethod = new ExtractedMethod();
        //            var extractedMethodModel = new ExtractMethodModel(extractedMethod, paramClassify);
        //            extractedMethodModel.extract(declarations, qs.Value, code);
        //            return extractedMethodModel;
        //        };

        //        var extraction = new ExtractMethodExtraction();
        //        // bug: access to disposed closure - todo: make ExtractMethodRefactoring request reparse like everyone else.
        //        Action<object> parseRequest = obj => _state.OnParseRequested(obj); 

        //        var refactoring = new ExtractMethodRefactoring(module, parseRequest, createMethodModel, extraction);
        //        refactoring.InvalidSelection += HandleInvalidSelection;
        //        refactoring.Refactor();
        //    }
        //    */
        //}
    }
}
