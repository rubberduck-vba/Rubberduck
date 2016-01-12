using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerCustomFolderViewModel : ViewModelBase
    {
        private readonly string _name;
        private readonly IEnumerable<CodeExplorerComponentViewModel> _components;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.Class, 
            DeclarationType.Document, 
            DeclarationType.Module, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerCustomFolderViewModel(string name, IEnumerable<Declaration> declarations)
        {
            _name = name;

            var items = declarations.ToList();

            _components = items.GroupBy(item => item.ComponentName)
                .SelectMany(grouping =>
                    grouping.Where(item => ComponentTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerComponentViewModel(item, grouping)))
                .OrderBy(item => item.Name)
                .ToList();
        }

        public string Name { get { return _name; } }
        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
    }
}