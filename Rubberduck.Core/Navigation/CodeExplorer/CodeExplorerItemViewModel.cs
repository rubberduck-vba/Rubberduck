using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigation.CodeExplorer
{
    public abstract class CodeExplorerItemViewModel : CodeExplorerItemViewModelBase
    {
        protected CodeExplorerItemViewModel(ICodeExplorerNode parent, Declaration declaration) : base(parent, declaration) { }

        public override string Name => Declaration?.IdentifierName ?? string.Empty;

        public override string NameWithSignature => Name;

        private bool _isErrorState;
        public override bool IsErrorState
        {
            get => _isErrorState;
            set
            {
                if (_isErrorState == value)
                {
                    return;
                }

                _isErrorState = value;

                foreach (var child in Children)
                {
                    child.IsErrorState = _isErrorState;
                }

                OnPropertyChanged();
            }
        }

        public virtual void Synchronize(List<Declaration> updated)
        {
            if (Declaration is null)
            {
                return;
            }

            var matching = updated.FirstOrDefault(decl => Declaration.DeclarationType == decl?.DeclarationType && Declaration.QualifiedName.Equals(decl?.QualifiedName));

            if (matching is null)
            {
                Declaration = null;
                return;
            }

            Declaration = matching;
            updated.Remove(matching);
            SynchronizeChildren(updated);
        }

        protected virtual void SynchronizeChildren(List<Declaration> updated)
        {
            foreach (var child in Children.OfType<CodeExplorerItemViewModel>().ToList())
            {
                child.Synchronize(updated);
                if (child.Declaration is null)
                {
                    RemoveChild(child);
                    continue;
                }

                updated.Remove(child.Declaration);
            }

            AddNewChildren(updated);
        }

        protected abstract void AddNewChildren(List<Declaration> updated);
    }
}
