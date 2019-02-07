using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Navigation.CodeExplorer
{
    public static class CodeExplorerItemComparer
    {
        public static Comparer<ICodeExplorerNode> NodeType { get; } = new CompareByNodeType();

        public static Comparer<ICodeExplorerNode> Name { get; } = new CompareByName();

        public static Comparer<ICodeExplorerNode> DeclarationType { get; } = new CompareByDeclarationType();
 
        public static Comparer<ICodeExplorerNode> Accessibility { get; } = new CompareByAccessibility();

        public static Comparer<ICodeExplorerNode> CodeLine { get; } = new CompareByCodeLine();

        public static Comparer<ICodeExplorerNode> ComponentType { get; } = new CompareByComponentType();

        public static Comparer<ICodeExplorerNode> DeclarationTypeThenName { get; } = new CompareByDeclarationTypeAndName();

        public static Comparer<ICodeExplorerNode> DeclarationTypeThenCodeLine { get; } = new CompareByDeclarationTypeAndCodeLine();

        public static Comparer<ICodeExplorerNode> ReferencePriority { get; } = new CompareByReferencePriority();

        public static Comparer<ICodeExplorerNode> ReferenceType { get; } = new CompareByReferenceType();
    }

    public class CompareByDeclarationTypeAndCodeLine : Comparer<ICodeExplorerNode>
    {
        private static readonly List<Func<ICodeExplorerNode, ICodeExplorerNode, int>> Comparisons =
            new List<Func<ICodeExplorerNode, ICodeExplorerNode, int>>
            {
                (x, y) => CodeExplorerItemComparer.DeclarationType.Compare(x, y),
                (x, y) => CodeExplorerItemComparer.CodeLine.Compare(x, y)
            };

        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            return x == y ? 0 : Comparisons.Select(comp => comp(x, y)).FirstOrDefault(result => result != 0);
        }
    }

    public class CompareByDeclarationTypeAndName : Comparer<ICodeExplorerNode>
    {
        private static readonly List<Func<ICodeExplorerNode, ICodeExplorerNode, int>> Comparisons =
            new List<Func<ICodeExplorerNode, ICodeExplorerNode, int>>
            {
                (x, y) => CodeExplorerItemComparer.DeclarationType.Compare(x, y),               
                (x, y) => CodeExplorerItemComparer.Name.Compare(x, y),
                (x, y) => CodeExplorerItemComparer.Accessibility.Compare(x, y)
            };

        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            return x == y ? 0 : Comparisons.Select(comp => comp(x, y)).FirstOrDefault(result => result != 0);
        }
    }

    public class CompareByName : CompareByNodeType
    {
        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            var node = base.Compare(x, y);
            return node != 0 ? node : string.Compare(x?.NameWithSignature, y?.NameWithSignature, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class CompareByCodeLine : CompareByNodeType
    {
        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            var node = base.Compare(x, y);
            if (node != 0)
            {
                return node;
            }

            var first = x?.QualifiedSelection?.Selection;

            if (first is null)
            {
                return -1;
            }

            var second = y?.QualifiedSelection?.Selection;

            if (second is null)
            {
                return 1;
            }

            return first.Value.CompareTo(second.Value);
        }
    }

    public class CompareByDeclarationType : Comparer<ICodeExplorerNode>
    {
        private static readonly Dictionary<DeclarationType, int> SortOrder = new Dictionary<DeclarationType, int>
        {
            // Some DeclarationTypes we want to treat the same, like Subs and Functions,
            // or Property Gets, Lets, and Sets.
            // Give them the same number.
            {DeclarationType.LibraryFunction, 0},
            {DeclarationType.LibraryProcedure, 0},
            {DeclarationType.UserDefinedType, 1},
            {DeclarationType.Enumeration, 2},
            {DeclarationType.Event, 3},
            {DeclarationType.Constant, 4},
            {DeclarationType.Variable, 5},
            {DeclarationType.PropertyGet, 6},
            {DeclarationType.PropertyLet, 6},
            {DeclarationType.PropertySet, 6},
            {DeclarationType.Function, 7},
            {DeclarationType.Procedure, 7}
        };

        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            if (x == y)
            {
                return 0;
            }

            var node = CodeExplorerItemComparer.NodeType.Compare(x, y);
            if (node != 0)
            {
                return node;
            }

            var first = x?.Declaration?.DeclarationType;

            if (first is null || !SortOrder.ContainsKey(first.Value))
            {
                return -1;
            }

            var second = y?.Declaration?.DeclarationType;

            if (second is null || !SortOrder.ContainsKey(second.Value))
            {
                return 1;
            }

            return SortOrder[first.Value].CompareTo(SortOrder[second.Value]);
        }
    }

    public class CompareByComponentType : Comparer<ICodeExplorerNode>
    {
        private static readonly Dictionary<ComponentType, int> SortOrder = new Dictionary<ComponentType, int>
        {
            // These are intended to be in the same order as the host would display them in folder view.
            // Not sure about some of the VB6 specific ones.
            {ComponentType.Document, 0},
            {ComponentType.UserForm, 1},
            {ComponentType.VBForm, 1},
            {ComponentType.MDIForm, 1},
            {ComponentType.StandardModule, 2},
            {ComponentType.ClassModule, 3},
            {ComponentType.UserControl, 4},
            {ComponentType.ActiveXDesigner, 4},
            {ComponentType.PropPage, 5},
            {ComponentType.DocObject, 6},
            {ComponentType.ResFile, 6},
            {ComponentType.RelatedDocument, 6},           
            {ComponentType.ComComponent, 7},
            {ComponentType.Undefined, 7}
        };

        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            if (x == y)
            {
                return 0;
            }

            var first = x?.QualifiedSelection?.QualifiedName.ComponentType;

            if (first is null || !SortOrder.ContainsKey(first.Value))
            {
                return -1;
            }

            var second = y?.QualifiedSelection?.QualifiedName.ComponentType;

            if (second is null || !SortOrder.ContainsKey(second.Value))
            {
                return 1;
            }

            var component = SortOrder[first.Value].CompareTo(SortOrder[second.Value]);

            return component == 0 ? CodeExplorerItemComparer.Name.Compare(x, y) : component;
        }
    }

    public class CompareByAccessibility : Comparer<ICodeExplorerNode>
    {
        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            if (x == y)
            {
                return 0;
            }

            var first = x?.Declaration?.Accessibility;

            if (first is null)
            {
                return -1;
            }

            var second = y?.Declaration?.Accessibility;

            if (second is null)
            {
                return 1;
            }

            // Public and Implicit Subs/Functions appear the same, so treat Implicits like Publics.
            if (first == Accessibility.Implicit)
            {
                first = Accessibility.Public;
            }

            if (second == Accessibility.Implicit)
            {
                first = Accessibility.Public;
            }

            // These are reversed because Accessibility is ordered lowest to highest.
            return second.Value.CompareTo(first.Value);
        }
    }

    public class CompareByNodeType : Comparer<ICodeExplorerNode>
    {
        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            if (x == y)
            {
                return 0;
            }

            if (x == null)
            {
                return -1;
            }

            if (y == null)
            {
                return 1;
            }

            // references come first
            if (x is CodeExplorerReferenceFolderViewModel ^
                y is CodeExplorerReferenceFolderViewModel)
            {
                return x is CodeExplorerReferenceFolderViewModel ? -1 : 1;
            }

            // folders come next
            if (x is CodeExplorerCustomFolderViewModel ^
                y is CodeExplorerCustomFolderViewModel)
            {
                return x is CodeExplorerCustomFolderViewModel ? -1 : 1;
            }

            return 0;
        }
    }

    public class CompareByReferencePriority : Comparer<ICodeExplorerNode>
    {
        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            if (x == y)
            {
                return 0;
            }

            if (!(x is CodeExplorerReferenceViewModel first))
            {
                return -1;
            }

            if (!(y is CodeExplorerReferenceViewModel second))
            {
                return 1;
            }

            return (first.Reference?.Priority ?? int.MaxValue).CompareTo(second.Reference?.Priority);
        }
    }

    public class CompareByReferenceType : CompareByNodeType
    {
        public override int Compare(ICodeExplorerNode x, ICodeExplorerNode y)
        {
            var node = base.Compare(x, y);
            if (node != 0)
            {
                return node;
            }

            if (!(x is CodeExplorerReferenceFolderViewModel first))
            {
                return -1;
            }

            if (!(y is CodeExplorerReferenceFolderViewModel second))
            {
                return 1;
            }

            // Libraries are first, so reverse the comparison.
            return second.ReferenceKind.CompareTo(first.ReferenceKind);
        }
    }
}
