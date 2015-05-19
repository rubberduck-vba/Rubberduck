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
            _declarations = parseResult.Declarations;
            _selection = selection;

            _view.OkButtonClicked += OnOkButtonClicked;
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
                        MessageBox.Show(RubberduckUI.ReorderParamsDialog_OptionalVariableError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            var indexOfParamArray = _view.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0)
            {
                if (indexOfParamArray != _view.Parameters.Count - 1)
                {
                    MessageBox.Show(RubberduckUI.ReorderParamsDialog_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            AdjustSignatures();
            AdjustReferences();
        }

        private void AdjustReferences()
        {
            foreach (var reference in _view.Target.References.Where(item => item.Context != _view.Target.Context))
            {
                var proc = (dynamic)reference.Context.Parent;
                var module = reference.QualifiedModuleName.Component.CodeModule;

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

                var argList = (VBAParser.ArgsCallContext)proc.argsCall();

                if (argList == null)
                {
                    continue;
                }

                RewriteCall(reference, argList, module);
            }
        }

        private void RewriteCall(IdentifierReference reference, VBAParser.ArgsCallContext argList, Microsoft.Vbe.Interop.CodeModule module)
        {
            var paramNames = argList.argCall().Select(arg => arg.GetText()).ToList();

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

        private void AdjustSignatures()
        {
            var proc = (dynamic)_view.Target.Context;
            var argList = (VBAParser.ArgListContext)proc.argList();
            var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (_view.Target.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _view.Target.ParentScope &&
                                              item.IdentifierName == _view.Target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter);
                }

                var letter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _view.Target.ParentScope &&
                              item.IdentifierName == _view.Target.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }

            RewriteSignature(argList, module);

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _view.Target.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_view.Target.Project) &&
                                                               item.IdentifierName == _view.Target.ComponentName + "_" + _view.Target.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustSignatures(interfaceImplentation);
            }
        }

        private void AdjustSignatures(IdentifierReference reference)
        {
            var proc = (dynamic)reference.Context.Parent;
            var module = reference.QualifiedModuleName.Component.CodeModule;
            var argList = (VBAParser.ArgListContext)proc.argList();

            RewriteSignature(argList, module);
        }

        private void AdjustSignatures(Declaration reference)
        {
            var proc = (dynamic)reference.Context.Parent;
            var module = reference.QualifiedName.QualifiedModuleName.Component.CodeModule;
            VBAParser.ArgListContext argList;

            if (reference.DeclarationType == DeclarationType.PropertySet || reference.DeclarationType == DeclarationType.PropertyLet)
            {
                argList = (VBAParser.ArgListContext)proc.children[0].argList();
            }
            else
            {
                argList = (VBAParser.ArgListContext)proc.subStmt().argList();
            }

            RewriteSignature(argList, module);
        }

        private void RewriteSignature(VBAParser.ArgListContext argList, Microsoft.Vbe.Interop.CodeModule module)
        {
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
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                     || IsSelectedReference(selection, item));

            while (target != null && !ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                target = _declarations.Items
                    .Where(item => item.QualifiedName.MemberName == target.ParentScope.Substring(target.ParentScope.LastIndexOf('.') + 1)
                                && item.Scope == target.ParentScope).FirstOrDefault();
            }

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
