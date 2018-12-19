using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CompareByName : Comparer<CodeExplorerItemViewModel>
    {
        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            var nodeComparison = new CompareByNodeType().Compare(x, y);

            return nodeComparison != 0
                ? nodeComparison
                : string.Compare(x?.NameWithSignature, y?.NameWithSignature, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class CompareByType : Comparer<CodeExplorerItemViewModel>
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

        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            var nodeComparison = new CompareByNodeType().Compare(x, y);
            if (nodeComparison != 0)
            {
                return nodeComparison;
            }

            //null sorts last.
            if (x?.Declaration == null)
            {
                return 1;
            }

            if (y?.Declaration == null)
            {
                return -1;
            }

            // keep separate types separate
            if (x.Declaration.DeclarationType != y.Declaration.DeclarationType)
            {
                if (SortOrder.TryGetValue(x.Declaration.DeclarationType, out var xValue) &&
                    SortOrder.TryGetValue(y.Declaration.DeclarationType, out var yValue))
                {
                    return xValue.CompareTo(yValue);
                }
            }

            // The Tree shows Public and Private Subs/Functions with a separate icon.
            // But Public and Implicit Subs/Functions appear the same, so treat Implicits like Publics.
            var xNodeAcc = x.Declaration.Accessibility == Accessibility.Implicit ? Accessibility.Public : x.Declaration.Accessibility;
            var yNodeAcc = y.Declaration.Accessibility == Accessibility.Implicit ? Accessibility.Public : y.Declaration.Accessibility;

            if (xNodeAcc != yNodeAcc)
            {
                // These are reversed because Accessibility is ordered lowest to highest.
                return yNodeAcc.CompareTo(xNodeAcc);
            }

            if (x.ExpandedIcon != y.ExpandedIcon)
            {
                // ReSharper disable PossibleInvalidOperationException - this will have a QualifiedSelection
                var xQmn = x.QualifiedSelection.Value.QualifiedName;
                var yQmn = y.QualifiedSelection.Value.QualifiedName;

                if (xQmn.ComponentType == ComponentType.Document ^ yQmn.ComponentType == ComponentType.Document)
                {
                    return xQmn.ComponentType == ComponentType.Document ? -1 : 1;
                }
            }

            return 0;
        }
    }

    public class CompareBySelection : Comparer<CodeExplorerItemViewModel>
    {
        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            var nodeComparison = new CompareByNodeType().Compare(x, y);
            if (nodeComparison != 0)
            {
                return nodeComparison;
            }

            // ReSharper disable PossibleNullReferenceException - tested in CompareByNodeType() above
            if (!x.QualifiedSelection.HasValue && !y.QualifiedSelection.HasValue)
               
            {
                return 0;
            }

            if (x.QualifiedSelection.HasValue ^ y.QualifiedSelection.HasValue)
            {
                return x.QualifiedSelection.HasValue ? -1 : 1;
            }

            if (x.QualifiedSelection.Value.Selection.StartLine == y.QualifiedSelection.Value.Selection.StartLine)
            {
                return 0;
            }

            return x.QualifiedSelection.Value.Selection.StartLine < y.QualifiedSelection.Value.Selection.StartLine ? -1 : 1;
            // ReSharper restore PossibleNullReferenceException
        }
    }

    public class CompareByNodeType : Comparer<CodeExplorerItemViewModel>
    {
        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            //null sorts last.
            if (x == null)
            {
                return 1;
            }

            if (y == null)
            {
                return -1;
            }

            // references come first
            if (x is CodeExplorerReferenceFolderViewModel ^
                y is CodeExplorerReferenceFolderViewModel)
            {
                return x is CodeExplorerReferenceFolderViewModel ? -1 : 1;
            }

            // references always sort by priority
            if (x is CodeExplorerReferenceViewModel first &&
                y is CodeExplorerReferenceViewModel second)
            {
                return first.Priority > second.Priority ? 1 : -1;
            }

            // folders come next
            if (x is CodeExplorerCustomFolderViewModel ^
                y is CodeExplorerCustomFolderViewModel)
            {
                return x is CodeExplorerCustomFolderViewModel ? -1 : 1;
            }

            // folders are always sorted by name
            if (x is CodeExplorerCustomFolderViewModel)
            {
                return string.CompareOrdinal(x.NameWithSignature, y.NameWithSignature);
            }

            return 0;
        }
    }

}
