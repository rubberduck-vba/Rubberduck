using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;
using System;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using NLog;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RefactorExtractMethodCommand(VBE vbe, RubberduckParserState state, IIndenter indenter)
            : base (vbe)
        {
            _state = state;
            _indenter = indenter;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var qualifiedSelection = Vbe.ActiveCodePane.GetQualifiedSelection();
            if (!qualifiedSelection.HasValue)
            {
                return false;
            }
            var selection = qualifiedSelection.Value.Selection;

            var code = Vbe.ActiveCodePane.CodeModule.Lines[selection.StartLine, selection.LineCount];

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

            _logger.Debug("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            var declarations = _state.AllDeclarations;
            var qualifiedSelection = Vbe.ActiveCodePane.GetQualifiedSelection();

            var extractMethodValidation = new ExtractMethodSelectionValidation(declarations);
            var canExecute = extractMethodValidation.withinSingleProcedure(qualifiedSelection.Value);
            if (!canExecute)
            {
                return;
            }
            ICodeModuleWrapper codeModuleWrapper = new CodeModuleWrapper(Vbe.ActiveCodePane.CodeModule);
            VBComponent vbComponent = Vbe.SelectedVBComponent;

            Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = ( qs, code) =>
            {
                if (qs == null)
                {
                    return null;
                }
                //TODO: Pull these even further back;
                //      and implement with IProvider<IExtractMethodRule>
                var rules = new List<IExtractMethodRule>(){ 
                    new ExtractMethodRuleInSelection(),
                    new ExtractMethodRuleIsAssignedInSelection(),
                    new ExtractMethodRuleUsedAfter(),
                    new ExtractMethodRuleUsedBefore()};

                var paramClassify = new ExtractMethodParameterClassification(rules);

                var extractedMethod = new ExtractedMethod();
                var extractedMethodModel = new ExtractMethodModel(extractedMethod,paramClassify);
                extractedMethodModel.extract(declarations, qs.Value, code);
                return extractedMethodModel;
            };

            var extraction = new ExtractMethodExtraction();
            Action<Object> parseRequest = (obj) => _state.OnParseRequested(obj, vbComponent);

            var refactoring = new ExtractMethodRefactoring(codeModuleWrapper, parseRequest , createMethodModel, extraction);
            refactoring.InvalidSelection += HandleInvalidSelection;
            refactoring.Refactor();
        }
    }
}
