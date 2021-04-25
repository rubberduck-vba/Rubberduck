using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public class ArgumentReference : IdentifierReference
    {
        internal ArgumentReference(
            QualifiedModuleName qualifiedName,
            Declaration parentScopingDeclaration,
            Declaration parentNonScopingDeclaration,
            string identifierName,
            Selection argumentSelection,
            ParserRuleContext context,
            VBAParser.ArgumentListContext argumentListContext,
            ArgumentListArgumentType argumentType,
            int argumentPosition,
            ParameterDeclaration referencedParameter,
            IEnumerable<IParseTreeAnnotation> annotations = null)
            : base(
                qualifiedName,
                parentScopingDeclaration,
                parentNonScopingDeclaration,
                identifierName,
                argumentSelection,
                context,
                referencedParameter,
                false,
                false,
                annotations)
        {
            ArgumentType = argumentType;
            ArgumentPosition = argumentPosition;
            ArgumentListContext = argumentListContext;
            NumberOfArguments = ArgumentListContext?.argument()?.Length ?? 0;
            ArgumentListSelection = argumentListContext?.GetSelection() ?? Selection.Empty;
        }

        public ArgumentListArgumentType ArgumentType { get; }
        public int ArgumentPosition { get; }
        public int NumberOfArguments { get; }
        public VBAParser.ArgumentListContext ArgumentListContext { get; }
        public Selection ArgumentListSelection { get; }

        public override (string context, Selection highlight) HighligthSelection(ICodeModule module)
        {
            const int maxLength = 255;

            var lines = module.GetLines(Selection.StartLine, Selection.LineCount).Split('\n');

            var line = lines[0];
            var indent = line.TakeWhile(c => c.Equals(' ')).Count();

            var highlight = new Selection(
                    1, Math.Max(Selection.StartColumn - indent, 1),
                    1, Math.Max(Selection.EndColumn - indent, 1))
                .ToZeroBased();

            var trimmed = line.Trim();
            if (trimmed.Length > maxLength || lines.Length > 1)
            {
                trimmed = trimmed.Substring(0, Math.Min(trimmed.Length, maxLength)) + "...";
                highlight = new Selection(1, highlight.StartColumn, 1, trimmed.Length);
            }

            if (highlight.IsSingleCharacter && highlight.StartColumn == 0)
            {
                trimmed = " " + trimmed;
                highlight = new Selection(0, 0, 0, 1);
            }
            else if (highlight.IsSingleCharacter)
            {
                highlight = new Selection(0, Selection.StartColumn-1 - indent - 1, 0, Selection.StartColumn-1 - indent);
            }
            return (trimmed, highlight);
        }

    }
}