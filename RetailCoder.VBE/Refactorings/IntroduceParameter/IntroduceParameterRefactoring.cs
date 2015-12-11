using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class PromoteLocalToParameterRefactoring : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly IActiveCodePaneEditor _editor;
        private Declaration _targetDeclaration;
        private IMessageBox _messageBox;

        public PromoteLocalToParameterRefactoring (RubberduckParserState parseResult, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _declarations = parseResult.AllDeclarations.ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor ()
        {
            if (_targetDeclaration == null)
            {
                _messageBox.Show("Invalid selection...");   // todo: write a better message and localize it
                return;
            }

            RemoveVariable();
        }

        public void Refactor (QualifiedSelection target)
        {
            _targetDeclaration = _declarations.FindSelection(target, new [] {DeclarationType.Variable});
            var v = _declarations.Where(i => !i.IsBuiltIn);

            _editor.SetSelection(target);
            Refactor();
        }

        public void Refactor (Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            _targetDeclaration = target;
            _editor.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        private void AddParameter ()
        {
            // insert string from GetParameterDefinition() into sub/function/... declaration
        }

        private void RemoveVariable()
        {
            var selection = new Selection(_targetDeclaration.Context.Start.Line,
                                          _targetDeclaration.Context.Start.Column,
                                          _targetDeclaration.Context.Stop.Line,
                                          _targetDeclaration.Context.Stop.Column);

            var lines =
                _editor.GetLines(selection)
                    .Replace(Environment.NewLine, string.Empty)
                    .Replace(" _", string.Empty)
                    .Remove(selection.StartColumn, _targetDeclaration.Context.GetText().Length);

            _editor.DeleteLines(selection);
            _editor.InsertLines(selection.StartLine, lines);
        }

        private string GetParameterDefinition ()
        {
            if (_targetDeclaration == null) { return null; }

            return _targetDeclaration.IdentifierName + " As " + _targetDeclaration.AsTypeName;
        }
    }
}
