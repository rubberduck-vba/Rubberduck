using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;

        public RefactorExtractMethodCommand(IVBE vbe, RubberduckParserState state, IIndenter indenter)
            : base (vbe)
        {
            _state = state;
            _indenter = indenter;
        }

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

                var allDeclarations = _state.AllDeclarations;
                var extractMethodValidation = new ExtractMethodSelectionValidation(allDeclarations);
                
                var canExecute = extractMethodValidation.withinSingleProcedure(qualifiedSelection.Value);

                return canExecute;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var declarations = _state.AllDeclarations;
            var qualifiedSelection = Vbe.ActiveCodePane.GetQualifiedSelection();

            var extractMethodValidation = new ExtractMethodSelectionValidation(declarations);
            var canExecute = extractMethodValidation.withinSingleProcedure(qualifiedSelection.Value);
            if (!canExecute)
            {
                return;
            }

            var pane = Vbe.ActiveCodePane;
            var module = pane.CodeModule;
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
    }
}
