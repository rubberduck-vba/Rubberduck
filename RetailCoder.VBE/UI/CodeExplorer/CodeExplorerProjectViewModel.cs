using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerProjectViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly IEnumerable<CodeExplorerComponentViewModel> _components;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.Class, 
            DeclarationType.Document, 
            DeclarationType.Module, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerProjectViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            _components = declarations.GroupBy(item => item.ComponentName)
                .SelectMany(grouping =>
                    grouping.Where(item => ComponentTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerComponentViewModel(item, grouping)))
                        .ToList();
        }

        public bool IsProtected { get { return _declaration.Project.Protection == vbext_ProjectProtection.vbext_pp_locked; } }
        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
    }
}