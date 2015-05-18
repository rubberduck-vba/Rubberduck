using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    class ReorderParametersPresenter
    {
        private readonly IReorderParametersView _view;
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;

        public ReorderParametersPresenter(IReorderParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _view = view;
            _view.OkButtonClicked += OnOkButtonClicked;

            _declarations = parseResult.Declarations;
            _selection = selection;
        }

        public void Show()
        {
            AcquireTarget(_selection);

            if (_view.Target != null)
            {
                LoadParameters();

                if (_view.Parameters.Count < 2) { return ;}

                _view.InitializeParameterGrid();
               _view.ShowDialog();
            }
        }

        private void LoadParameters()
        {
            var procedure = (dynamic)_view.Target.Context;
            var argList = (VBAParser.ArgListContext)procedure.argList();
            var args = argList.arg();

            var index = 0;
            foreach (var arg in args)
            {
                _view.Parameters.Add(new Parameter(arg.GetText(), index++));
            }
        }

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            if (!_view.Parameters.Where((t, i) => t.Index != i).Any())
            {
                return;
            }

            var indexOfFirstOptionalParam = _view.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < _view.Parameters.Count; index++)
                {
                    if (!_view.Parameters.ElementAt(index).IsOptional)
                    {
                        var message = "Optional parameters must be specified at the end of the parameter list.";
                        MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            AdjustSignature();
            AdjustReferences();

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _view.Target.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustSignature(reference);
                }
            }

            if (_view.Target.DeclarationType == DeclarationType.Procedure)
            {
                foreach (var interfaceImplentation in
                    _declarations.FindInterfaceImplementationMembers()
                                 .Where(item => item.Project.Equals(_view.Target.Project) && item.IdentifierName.Contains(_view.Target.ComponentName)))
                {
                    AdjustSignature(interfaceImplentation);
                }
            }
        }

        private void AdjustReferences()
        {
            foreach (var reference in _view.Target.References.Where(item => item.Context != _view.Target.Context))
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
                    // update letter methods - needs proper fixing
                    if (reference.Context.Parent.GetText().Contains("Property Let"))
                    {
                        AdjustSignature(reference);
                    }

                    continue;
                }

                var argList = (VBAParser.ArgsCallContext)proc.argsCall();

                if (argList == null)
                {
                    continue;
                }
                var paramNames = argList.argCall().Select(arg => arg.GetText()).ToList();

                var module = reference.QualifiedModuleName.Component.CodeModule;
                var lineCount = argList.Stop.Line - argList.Start.Line + 1; // adjust for total line count

                var variableIndex = 0;
                for (var line = argList.Start.Line; line < argList.Start.Line + lineCount; line++)
                {
                    var newContent = module.Lines[line, 1].Replace(" , ", "");

                    var currentStringIndex = line == argList.Start.Line ? reference.Declaration.IdentifierName.Length : 0;

                    for (var i = 0; i < paramNames.Count && variableIndex < _view.Parameters.Count; i++)
                    {
                        var variableStringIndex = newContent.IndexOf(paramNames.ElementAt(i), currentStringIndex);

                        if (variableStringIndex > -1)
                        {
                            if (_view.Parameters.ElementAt(variableIndex).Index >= paramNames.Count)
                            {
                                newContent = newContent.Insert(variableStringIndex, " , ");
                                i--;
                                variableIndex++;
                                continue;
                            }

                            var oldVariableString = paramNames.ElementAt(i);
                            var newVariableString = paramNames.ElementAt(_view.Parameters.ElementAt(variableIndex).Index);
                            var beginningSub = newContent.Substring(0, variableStringIndex);
                            var replaceSub = newContent.Substring(variableStringIndex).Replace(oldVariableString, newVariableString);

                            newContent = beginningSub + replaceSub;

                            variableIndex++;
                            currentStringIndex = beginningSub.Length + newVariableString.Length;
                        }
                    }

                    module.ReplaceLine(line, newContent);
                }
            }
        }

        // TODO - refactor this
        // Only used for Property Letters/Setters
        // Otherwise, they are caught by the try/catch block used to prevent 
        // value returns from crashing the program
        // Extremely similar to the other AdjustReference
        // One possibility is to create multiple "handler" methods
        // that call another method to do the real work, passing 
        // "module", "argList", and "args" (I believe these are the only
        // ones used
        private void AdjustSignature(IdentifierReference reference)
        {
            var proc = (dynamic)reference.Context.Parent;
            var module = reference.QualifiedModuleName.Component.CodeModule;
            var argList = (VBAParser.ArgListContext)proc.argList();
            var args = argList.arg();

            var variableIndex = 0;
            for (var lineNum = argList.Start.Line; lineNum < argList.Start.Line + argList.GetSelection().LineCount; lineNum++)
            {
                var newContent = module.Lines[lineNum, 1];
                var currentStringIndex = 0;

                for (var i = variableIndex; i < _view.Parameters.Count; i++)
                {
                    var variableStringIndex = newContent.IndexOf(_view.Parameters.Find(item => item.Index == variableIndex).FullDeclaration, currentStringIndex);

                    if (variableStringIndex > -1)
                    {
                        var oldVariableString = _view.Parameters.Find(item => item.Index == variableIndex).FullDeclaration;
                        var newVariableString = _view.Parameters.ElementAt(i).FullDeclaration;
                        var beginningSub = newContent.Substring(0, variableStringIndex);
                        var replaceSub = newContent.Substring(variableStringIndex).Replace(oldVariableString, newVariableString);

                        newContent = beginningSub + replaceSub;

                        variableIndex++;
                        currentStringIndex = beginningSub.Length + newVariableString.Length;
                    }
                }

                module.ReplaceLine(lineNum, newContent);
            }
        }

        private void AdjustSignature(Declaration reference = null)
        {
            var proc = (dynamic)_view.Target.Context;
            var argList = (VBAParser.ArgListContext)proc.argList();
            var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;

            if (reference != null)
            {
                proc = (dynamic)reference.Context.Parent;
                module = reference.QualifiedName.QualifiedModuleName.Component.CodeModule;

                if (reference.DeclarationType == DeclarationType.PropertySet)
                {
                    argList = (VBAParser.ArgListContext)proc.children[0].argList();
                }
                else
                {
                    argList = (VBAParser.ArgListContext)proc.subStmt().argList();
                }
            }

            var args = argList.arg();

            // if we are reordering a property getter, check if we need to reorder a setter too
            // only check if the passed reference is null, otherwise we recursively check and have an SO
            if (reference == null && _view.Target.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _view.Target.ParentScope &&
                                              item.IdentifierName == _view.Target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignature(setter);
                }
            }

            var variableIndex = 0;
            for (var lineNum = argList.Start.Line; lineNum < argList.Start.Line + argList.GetSelection().LineCount; lineNum++)
            {
                var newContent = module.Lines[lineNum, 1];
                var currentStringIndex = 0;

                for (var i = variableIndex; i < _view.Parameters.Count; i++)
                {
                    var variableStringIndex = newContent.IndexOf(_view.Parameters.Find(item => item.Index == variableIndex).FullDeclaration, currentStringIndex);

                    if (variableStringIndex > -1)
                    {
                        var oldVariableString = _view.Parameters.Find(item => item.Index == variableIndex).FullDeclaration;
                        var newVariableString = _view.Parameters.ElementAt(i).FullDeclaration;
                        var beginningSub = newContent.Substring(0, variableStringIndex);
                        var replaceSub = newContent.Substring(variableStringIndex).Replace(oldVariableString, newVariableString);

                        newContent = beginningSub + replaceSub;

                        variableIndex++;
                        currentStringIndex = beginningSub.Length + newVariableString.Length;
                    }
                }

                module.ReplaceLine(lineNum, newContent);
            }
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

        private void AcquireTarget(QualifiedSelection selection)
        {
            var target = _declarations.Items
                .Where(item => !item.IsBuiltIn && ValidDeclarationTypes.Contains(item.DeclarationType))
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                     || IsSelectedReference(selection, item));

            if (target != null && target.DeclarationType == DeclarationType.PropertySet)
            {
                var getter = _declarations.Items.FirstOrDefault(item => item.ParentScope == target.ParentScope &&
                                              item.IdentifierName == target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertyGet);

                if (getter != null)
                {
                    target = getter;
                }
            }

            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
        }

        private void PromptIfTargetImplementsInterface(ref Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (target == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.ReorderPresenter_TargetIsInterfaceMemberImplementation, target.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                target = null;
                return;
            }

            target = interfaceMember;
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
