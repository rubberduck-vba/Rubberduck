using System.Collections.Generic;
using System.Linq;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerReferenceFolderViewModel : CodeExplorerItemViewModelBase
    {
        private readonly DeclarationFinder _finder;

        public CodeExplorerReferenceFolderViewModel(
            ICodeExplorerNode parent, 
            DeclarationFinder finder, 
            List<ReferenceModel> references, 
            ReferenceKind type) 
            : base(parent, parent?.Declaration)
        {
            _finder = finder;
            ReferenceKind = type;
            Synchronize(Declaration, references);
        }

        public ReferenceKind ReferenceKind { get; }

        public override string Name => ReferenceKind == ReferenceKind.TypeLibrary
            ? Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_LibraryReferences
            : Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_ProjectReferences;

        public override string NameWithSignature => Name;

        public override string PanelTitle => Name;

        public override string Description => string.Empty;

        public override QualifiedSelection? QualifiedSelection => null;

        public override bool IsErrorState
        {
            get => false;
            set { }
        }

        public override bool Filtered => false;

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.ReferenceType;

        public void Synchronize(Declaration parent, List<ReferenceModel> updated)
        {
            var updates = updated.Where(reference => reference.Type == ReferenceKind).ToList();
            if (!updates.Any())
            {
                Declaration = null;
                return;
            }

            Declaration = parent;

            foreach (var child in Children.OfType<CodeExplorerReferenceViewModel>().ToList())
            {
                child.Synchronize(Declaration, updates);
                if (child.Reference is null)
                {
                    RemoveChild(child);
                    continue;
                }

                updated.Remove(child.Reference);
            }

            foreach (var reference in updates)
            {
                reference.IsUsed = reference.IsBuiltIn ||
                                   _finder != null &&
                                   _finder.IsReferenceUsedInProject(Declaration as ProjectDeclaration,
                                       reference.ToReferenceInfo());

                AddChild(new CodeExplorerReferenceViewModel(this, reference));
                updated.Remove(reference);
            }

            if (!Children.Any())
            {
                Declaration = null;
            }
        }

        public void UpdateChildren()
        {
            foreach (var library in Children.OfType<CodeExplorerReferenceViewModel>())
            {
                var reference = library.Reference;
                if (reference == null)
                {
                    continue;
                }

                reference.IsUsed = reference.IsBuiltIn ||
                                   _finder != null &&
                                   _finder.IsReferenceUsedInProject(
                                       library.Parent?.Declaration as ProjectDeclaration,
                                       reference.ToReferenceInfo());

                library.IsDimmed = !reference.IsUsed;
            }
        }
    }
}
