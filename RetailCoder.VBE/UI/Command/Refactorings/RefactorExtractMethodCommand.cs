using System.Diagnostics;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;
using System;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;

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

            var parentProcedure = _state.AllDeclarations.FindSelectedDeclaration(qualifiedSelection.Value, DeclarationExtensions.ProcedureTypes, d => ((ParserRuleContext)d.Context.Parent).GetSelection());
            var canExecute = parentProcedure != null
                && selection.StartColumn != selection.EndColumn
                && selection.LineCount > 0
                && !string.IsNullOrWhiteSpace(code);

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            var declarations = _state.AllDeclarations;
            ICodeModuleWrapper codeModuleWrapper = new CodeModuleWrapper(Vbe.ActiveCodePane.CodeModule);
            var qualifiedSelection = Vbe.ActiveCodePane.GetQualifiedSelection();

            Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = ( qs, code) =>
            {
                if (qs == null)
                {
                    return null;
                }

                return new ExtractMethodModel(declarations, qs.Value, code);
            };

            var createProc = new ExtractMethodProc();
            Func<IExtractMethodModel, string> createProcFunc = (model) => { return createProc.createProc(model); };
            var refactoring = new ExtractMethodRefactoring(codeModuleWrapper,createMethodModel, createProc);
            refactoring.InvalidSelection += HandleInvalidSelection;
            refactoring.Refactor();
        }
    }
}