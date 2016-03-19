using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerProjectViewModel : CodeExplorerItemViewModel
    {
        private readonly Declaration _declaration;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.Class, 
            DeclarationType.Document, 
            DeclarationType.Module, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerProjectViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            try
            {
                _declaration = declaration;
                Items = FindFolders(declarations.ToList(), '.');

                _icon = _declaration.Project.Protection == vbext_ProjectProtection.vbext_pp_locked
                    ? GetImageSource(resx.lock__exclamation)
                    : GetImageSource(resx.VSObject_Library);
            }
            catch (NullReferenceException e)
            {
                Console.WriteLine(e);
            }
        }

        private static IEnumerable<CodeExplorerItemViewModel> FindFolders(IEnumerable<Declaration> declarations, char delimiter)
        {
            var root = new CodeExplorerCustomFolderViewModel(string.Empty, new List<Declaration>());

            var items = declarations.ToList();
            var folders = items.Where(item => ComponentTypes.Contains(item.DeclarationType))
                               .GroupBy(item => item.CustomFolder)
                               .OrderBy(item => item.Key);
            foreach (var grouping in folders)
            {
                CodeExplorerItemViewModel node = root;
                var parts = grouping.Key.Split(delimiter);
                var path = new StringBuilder();
                foreach (var part in parts)
                {
                    if (path.Length != 0)
                    {
                        path.Append(delimiter);
                    }

                    path.Append(part);
                    var next = node.GetChild(part);
                    if (next == null)
                    {
                        var currentPath = path.ToString();
                        var parents = grouping.Where(item => ComponentTypes.Contains(item.DeclarationType) && item.CustomFolder == currentPath).ToList();

                        next = new CodeExplorerCustomFolderViewModel(part, items.Where(item => 
                            parents.Contains(item) || parents.Any(parent => 
                                (item.ParentDeclaration != null && item.ParentDeclaration.Equals(parent)) || item.ComponentName == parent.ComponentName)));
                        node.AddChild(next);
                    }

                    node = next;
                }
            }

            return root.Items;
        }

        private readonly BitmapImage _icon;
        public override BitmapImage CollapsedIcon { get { return _icon; } }
        public override BitmapImage ExpandedIcon { get { return _icon; } }

        public override string Name { get { return _declaration.CustomFolder; } }
        public override QualifiedSelection? QualifiedSelection { get { return _declaration.QualifiedSelection; } }
    }
}