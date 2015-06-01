using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    class RemoveParametersPresenter
    {
        private readonly IRemoveParametersView _view;
        private readonly VBProjectParseResult _parseResult;
        private readonly Declarations _declarations;
        private readonly Declaration _targetDeclaration;

        public RemoveParametersPresenter(IRemoveParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _view = view;
            _parseResult = parseResult;
            _declarations = parseResult.Declarations;

            FindTarget(out _targetDeclaration, selection);

            _view.OkButtonClicked += OkButtonClicked;
        }

        public void Show()
        {
            if (_targetDeclaration == null) { return; }

            _view.Parameters = LoadParameters();

            if (_view.Parameters.Count == 0)
            {
                var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, _targetDeclaration.IdentifierName);
                MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            _view.InitializeParameterGrid();
            _view.ShowDialog();
        }

        private List<Parameter> LoadParameters()
        {
            var index = 0;
            return FindTargets(_targetDeclaration).Select(arg => new Parameter(arg, index++)).ToList();
        }

        private void OkButtonClicked(object sender, EventArgs e)
        {
            var RemoveParams = new RemoveParameterRefactoring(_parseResult, _targetDeclaration, _view.Parameters);
            RemoveParams.Refactor();
        }

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private IEnumerable<Declaration> FindTargets(Declaration method)
        {
            return _declarations.Items
                              .Where(d => d.DeclarationType == DeclarationType.Parameter
                                       && d.ComponentName == method.ComponentName
                                       && d.Project.Equals(method.Project)
                                       && method.Context.Start.Line <= d.Selection.StartLine
                                       && method.Context.Stop.Line >= d.Selection.EndLine
                                       && !(method.Context.Start.Column > d.Selection.StartColumn && method.Context.Start.Line == d.Selection.StartLine)
                                       && !(method.Context.Stop.Column < d.Selection.EndColumn && method.Context.Stop.Line == d.Selection.EndLine))
                              .OrderBy(item => item.Selection.StartLine)
                              .ThenBy(item => item.Selection.StartColumn);
        }

        private void FindTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations.Items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                     || IsSelectedReference(selection, item));

            if (target != null && ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                return;
            }

            target = null;

            var targets = _declarations.Items
                .Where(item => !item.IsBuiltIn
                            && item.ComponentName == selection.QualifiedName.ComponentName
                            && ValidDeclarationTypes.Contains(item.DeclarationType));

            var currentStartLine = 0;
            var currentEndLine = int.MaxValue;
            var currentStartColumn = 0;
            var currentEndColumn = int.MaxValue;

            foreach (var declaration in targets)
            {
                var startLine = declaration.Context.Start.Line;
                var startColumn = declaration.Context.Start.Column;
                var endLine = declaration.Context.Stop.Line;
                var endColumn = declaration.Context.Stop.Column;

                if (startLine <= selection.Selection.StartLine && endLine >= selection.Selection.EndLine &&
                    currentStartLine <= startLine && currentEndLine >= endLine)
                {
                    if (!(startLine == selection.Selection.StartLine && (startColumn > selection.Selection.StartColumn || currentStartColumn > startColumn) ||
                          endLine == selection.Selection.EndLine && (endColumn < selection.Selection.EndColumn || currentEndColumn < endColumn)))
                    {
                        target = declaration;

                        currentStartLine = startLine;
                        currentEndLine = endLine;
                        currentStartColumn = startColumn;
                        currentEndColumn = endColumn;
                    }
                }

                foreach (var reference in declaration.References)
                {
                    var proc = (dynamic)reference.Context.Parent;

                    // This is to prevent throws when this statement fails:
                    // (VBAParser.ArgsCallContext)proc.argsCall();
                    try
                    {
                        var check = (VBAParser.ArgsCallContext)proc.argsCall();
                    }
                    catch
                    {
                        continue;
                    }

                    var paramList = (VBAParser.ArgsCallContext)proc.argsCall();

                    if (paramList == null)
                    {
                        continue;
                    }

                    startLine = paramList.Start.Line;
                    startColumn = paramList.Start.Column;
                    endLine = paramList.Stop.Line;
                    endColumn = paramList.Stop.Column + paramList.Stop.Text.Length + 1;

                    if (startLine <= selection.Selection.StartLine && endLine >= selection.Selection.EndLine &&
                        currentStartLine <= startLine && currentEndLine >= endLine)
                    {
                        if (!(startLine == selection.Selection.StartLine && (startColumn > selection.Selection.StartColumn || currentStartColumn > startColumn) ||
                              endLine == selection.Selection.EndLine && (endColumn < selection.Selection.EndColumn || currentEndColumn < endColumn)))
                        {
                            target = reference.Declaration;

                            currentStartLine = startLine;
                            currentEndLine = endLine;
                            currentStartColumn = startColumn;
                            currentEndColumn = endColumn;
                        }
                    }
                }
            }
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.References.Any(r =>
                r.QualifiedModuleName == selection.QualifiedName &&
                r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
