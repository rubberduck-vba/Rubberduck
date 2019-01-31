using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public interface ICodeExplorerNode
    {
        Declaration Declaration { get; }
        ICodeExplorerNode Parent { get; }
        ObservableCollection<ICodeExplorerNode> Children { get; }

        string Name { get; }
        string NameWithSignature { get; }
        string PanelTitle { get; }
        string Description { get; }

        QualifiedSelection? QualifiedSelection { get; }

        bool IsExpanded { get; set; }
        bool IsSelected { get; set; }
        bool IsDimmed { get; set; }
        bool IsObsolete { get; }
        bool IsErrorState { get; set; }
        string ToolTip { get; }
        FontWeight FontWeight { get; }

        CodeExplorerSortOrder SortOrder { get; set; }
        Comparer<ICodeExplorerNode> SortComparer { get; }
        string Filter { get; set; }
        bool Filtered { get; }
    }
}
