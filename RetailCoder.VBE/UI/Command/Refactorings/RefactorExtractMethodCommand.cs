using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using System;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Settings;
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

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.RefactorExtractMethod; }
        }

        protected override bool CanExecuteImpl(object parameter)
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
                var selection = qualifiedSelection.Value.Selection;

                var code = module.GetLines(selection.StartLine, selection.LineCount);

                var allDeclarations = _state.AllDeclarations;
                var extractMethodValidation = new ExtractMethodSelectionValidation(allDeclarations);
                //var parentProcedure = _state.AllDeclarations.FindSelectedDeclaration(qualifiedSelection.Value, DeclarationExtensions.ProcedureTypes, d => ((ParserRuleContext)d.Context.Parent).GetSelection());
                var canExecute = extractMethodValidation.withinSingleProcedure(qualifiedSelection.Value);

                /*
                var canExecute = parentProcedure != null
                    && selection.StartColumn != selection.EndColumn
                    && selection.LineCount > 0
                    && !string.IsNullOrWhiteSpace(code);
                */

                return canExecute;
            }
        }

        protected override void ExecuteImpl(object parameter)
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
            var component = module.Parent;
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
                Action<object> parseRequest = obj => _state.OnParseRequested(obj, component); 

                var refactoring = new ExtractMethodRefactoring(module, parseRequest, createMethodModel, extraction);
                refactoring.InvalidSelection += HandleInvalidSelection;
                refactoring.Refactor();
            }
        }
    }
}
