using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.PromoteLocalToParameter
{
    public class PromoteLocalToParameterRefactoring : IRefactoring
    {
        private readonly IActiveCodePaneEditor _editor;

        public void Refactor ()
        {
            GetParameterDefinition(_editor.GetSelection());
        }

        public void Refactor (QualifiedSelection target)
        {
            _editor.SetSelection(target);
            Refactor();
        }

        public void Refactor (Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            _editor.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        private void AddParameterToDefinition(Declaration target)
        {
            
        }

        private string GetParameterDefinition(QualifiedSelection? target)
        {
            if (target != null)
            {
                _editor.InsertLines(2, target.Value.QualifiedName.ComponentName);
            }

            return null;
        }
    }
}
