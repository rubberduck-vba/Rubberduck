using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel
    {
        private readonly RubberduckParserState _parseResult;
        public RubberduckParserState ParseResult { get { return _parseResult; } }

        private readonly IList<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        public Declaration TargetDeclaration { get; private set; }

        private readonly IMessageBox _messageBox;

        public EncapsulateFieldModel(RubberduckParserState parseResult, QualifiedSelection selection, IMessageBox messageBox)
        {
            _parseResult = parseResult;
            _declarations = parseResult.AllDeclarations.Where(d => !d.IsBuiltIn && d.DeclarationType == DeclarationType.Variable).ToList();
            _messageBox = messageBox;

            TargetDeclaration = FindSelection(selection);
        }

        private Selection GetVariableStmtContextSelection(Declaration target)
        {
            var statement = GetVariableStmtContext(target);

            return new Selection(statement.Start.Line, statement.Start.Column,
                    statement.Stop.Line, statement.Stop.Column);
        }

        private VBAParser.VariableStmtContext GetVariableStmtContext(Declaration target)
        {
            var statement = target.Context.Parent.Parent as VBAParser.VariableStmtContext;
            if (statement == null)
            {
                throw new NullReferenceException("Statement not found");
            }

            return statement;
        }

        private bool HasMultipleDeclarationsInStatement(Declaration target)
        {
            var statement = target.Context.Parent as VBAParser.VariableListStmtContext;

            return statement != null && statement.children.Count(i => i is VBAParser.VariableSubStmtContext) > 1;
        }

        private Declaration FindSelection(QualifiedSelection selection)
        {
            var target = _declarations
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));

            if (target != null) { return target; }

            var targets = _declarations.Where(item => item.ComponentName == selection.QualifiedName.ComponentName);

            foreach (var declaration in targets)
            {
                var declarationSelection = new Selection(declaration.Context.Start.Line,
                                                    declaration.Context.Start.Column,
                                                    declaration.Context.Stop.Line,
                                                    declaration.Context.Stop.Column + declaration.Context.Stop.Text.Length);

                if (declarationSelection.Contains(selection.Selection) ||
                    !HasMultipleDeclarationsInStatement(declaration) && GetVariableStmtContextSelection(declaration).Contains(selection.Selection))
                {
                    return declaration;
                }

                var reference =
                    declaration.References.FirstOrDefault(r => r.Selection.Contains(selection.Selection));

                if (reference != null)
                {
                    return reference.Declaration;
                }
            }
            return null;
        }
    }
}