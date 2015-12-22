using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerViewModelBuilder
    {
        private readonly RubberduckParserState _state;

        public CodeExplorerViewModelBuilder(RubberduckParserState state)
        {
            _state = state;
        }

        public IEnumerable<ExplorerItemViewModel> Build()
        {
            var userDeclarations = _state.AllDeclarations.Where(d => !d.IsBuiltIn).ToList();
            foreach (var projectDeclaration in userDeclarations.Where(d => d.DeclarationType == DeclarationType.Project))
            {
                var project = projectDeclaration;
                var projectItem = new ExplorerItemViewModel(project);
                foreach (var componentDeclaration in userDeclarations.Where(d => ReferenceEquals(d.ParentDeclaration, project)))
                {
                    var component = componentDeclaration;
                    yield return new ExplorerItemViewModel(component);
                    foreach (var member in userDeclarations.Where(d => ReferenceEquals(d.ParentDeclaration, component)))
                    {
                        yield return new ExplorerItemViewModel(member);
                        if (member.DeclarationType == DeclarationType.UserDefinedType)
                        {
                            
                        }

                        if (member.DeclarationType == DeclarationType.Enumeration)
                        {
                            
                        }
                    }
                }
            }
        }
    }
}