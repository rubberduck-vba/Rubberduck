using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Inspections
{
    public class IdentifierNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public IdentifierNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveUnusedDeclarationQuickFix(context, QualifiedSelection, Target), 
                new IgnoreOnceQuickFix(context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
        public override string Description 
        {
            get
            {
                return string.Format(InspectionsUI.IdentifierNotUsedInspectionResultFormat, Target.DeclarationType.ToLocalizedString(), Target.IdentifierName);
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }

    /// <summary>
    /// A code inspection quickfix that removes an unused identifier declaration.
    /// </summary>
    public class RemoveUnusedDeclarationQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;

        public RemoveUnusedDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target)
            : base(context, selection, InspectionsUI.RemoveUnusedDeclarationQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            if (_target.DeclarationType == DeclarationType.Variable)
            {
                RemoveVariable(_target);
            }
            else
            {
                var module = Selection.QualifiedName.Component.CodeModule;
                var selection = Selection.Selection;

                var originalCodeLines = module.Lines[selection.StartLine, selection.LineCount]
                    .Replace("\r\n", " ")
                    .Replace("_", string.Empty);

                var originalInstruction = Context.GetText();
                module.DeleteLines(selection.StartLine, selection.LineCount);

                var newCodeLines = originalCodeLines.Replace(originalInstruction, string.Empty);
                if (!string.IsNullOrEmpty(newCodeLines))
                {
                    module.InsertLines(selection.StartLine, newCodeLines);
                }
            }
        }

        private void RemoveVariable(Declaration target)
        {
            Selection selection;
            var declarationText = target.Context.GetText().Replace(" _" + Environment.NewLine, string.Empty);
            var multipleDeclarations = target.HasMultipleDeclarationsInStatement();

            var variableStmtContext = target.GetVariableStmtContext();

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText().Replace(" _" + Environment.NewLine, string.Empty);
                selection = target.GetVariableStmtContextSelection();
            }
            else
            {
                selection = new Selection(target.Context.Start.Line, target.Context.Start.Column,
                    target.Context.Stop.Line, target.Context.Stop.Column);
            }

            var codeModule = target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            var oldLines = codeModule.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(codeModule.GetLines(selection).Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            var newLinesWithoutExcessSpaces = newLines.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (var i = 0; i < newLinesWithoutExcessSpaces.Length; i++)
            {
                newLinesWithoutExcessSpaces[i] = newLinesWithoutExcessSpaces[i].RemoveExtraSpacesLeavingIndentation();
            }

            for (var i = newLinesWithoutExcessSpaces.Length - 1; i >= 0; i--)
            {
                if (newLinesWithoutExcessSpaces[i].Trim() == string.Empty)
                {
                    continue;
                }

                if (newLinesWithoutExcessSpaces[i].EndsWith(" _"))
                {
                    newLinesWithoutExcessSpaces[i] =
                        newLinesWithoutExcessSpaces[i].Remove(newLinesWithoutExcessSpaces[i].Length - 2);
                }
                break;
            }

            // remove all lines with only whitespace
            newLinesWithoutExcessSpaces = newLinesWithoutExcessSpaces.Where(str => str.Any(c => !char.IsWhiteSpace(c))).ToArray();

            codeModule.DeleteLines(selection);
            if (newLinesWithoutExcessSpaces.Any())
            {
                codeModule.InsertLines(selection.StartLine,
                    string.Join(Environment.NewLine, newLinesWithoutExcessSpaces));
            }
        }

        private string RemoveExtraComma(string str, int numParams, int indexRemoved)
        {
            // Example use cases for this method (fields and variables):
            // Dim fizz as Boolean, dizz as Double
            // Private fizz as Boolean, dizz as Double
            // Public fizz as Boolean, _
            //        dizz as Double
            // Private fizz as Boolean _
            //         , dizz as Double _
            //         , iizz as Integer

            // Before this method is called, the parameter to be removed has 
            // already been removed.  This means 'str' will look like:
            // Dim fizz as Boolean, 
            // Private , dizz as Double
            // Public fizz as Boolean, _
            //        
            // Private  _
            //         , dizz as Double _
            //         , iizz as Integer

            // This method is responsible for removing the redundant comma
            // and returning a string similar to:
            // Dim fizz as Boolean
            // Private dizz as Double
            // Public fizz as Boolean _
            //        
            // Private  _
            //          dizz as Double _
            //         , iizz as Integer

            var commaToRemove = numParams == indexRemoved ? indexRemoved - 1 : indexRemoved;

            return str.Remove(str.NthIndexOf(',', commaToRemove), 1);
        }
    }
}
