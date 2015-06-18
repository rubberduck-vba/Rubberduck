using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class VariableTypeNotDeclaredInspectionResult : CodeInspectionResultBase
    {
        public VariableTypeNotDeclaredInspectionResult(string inspection, CodeInspectionSeverity type, ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, type, qualifiedName, context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return
                new Dictionary<string, Action>
                {
                    {RubberduckUI.Inspections_DeclareAsExplicitVariant, DeclareAsExplicitVariant}
                };
        }

        private void DeclareAsExplicitVariant()
        {
            var codeModule = QualifiedSelection.QualifiedName.Component.CodeModule;
            var codeLine = codeModule.Lines[QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.LineCount];

            var context = (Context is VBAParser.VariableSubStmtContext ||
                           Context is VBAParser.ConstSubStmtContext ||
                           Context is VBAParser.ArgContext)
                          ? Context
                          : Context.Parent;

            // methods return empty string if soft-cast context is null - just concat results:
            string originalInstruction;
            var fix = DeclareExplicitVariant(context as VBAParser.VariableSubStmtContext, out originalInstruction);

            if (string.IsNullOrEmpty(originalInstruction))
            {
                fix = DeclareExplicitVariant(context as VBAParser.ConstSubStmtContext, out originalInstruction);
            }

            if (string.IsNullOrEmpty(originalInstruction))
            {
                fix = DeclareExplicitVariant(context as VBAParser.ArgContext, out originalInstruction);
            }
            
            var fixedCodeLine = codeLine.Replace(originalInstruction, fix);
            codeModule.ReplaceLine(QualifiedSelection.Selection.StartLine, fixedCodeLine);
        }

        private string DeclareExplicitVariant(VBAParser.VariableSubStmtContext context, out string instruction)
        {
            if (context == null)
            {
                instruction = null;
                return null;
            }

            instruction = context.GetText();
            return instruction + ' ' + Tokens.As + ' ' + Tokens.Variant;
        }

        private string DeclareExplicitVariant(VBAParser.ArgContext context, out string instruction)
        {
            if (context == null)
            {
                instruction = null;
                return null;
            }

            instruction = context.GetText();
            return instruction + ' ' + Tokens.As + ' ' + Tokens.Variant;
        }

        private string DeclareExplicitVariant(VBAParser.ConstSubStmtContext context, out string instruction)
        {
            if (context == null)
            {
                instruction = null;
                return null;
            }

            var parent = (VBAParser.ConstStmtContext) context.Parent;
            instruction = parent.GetText();

            var constant = context.GetText();
            var replacement = context.ambiguousIdentifier().GetText() + ' '
                         + Tokens.As + ' ' + Tokens.Variant + ' '
                         + context.EQ().GetText() + ' '
                         + context.valueStmt().GetText();

            var result = instruction.Replace(constant, replacement);
            return result;
        }
    }
}