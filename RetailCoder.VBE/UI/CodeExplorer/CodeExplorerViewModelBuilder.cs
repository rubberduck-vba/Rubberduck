using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
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
                Parallel.ForEach(userDeclarations.Where(d => ReferenceEquals(d.ParentDeclaration, project)), (componentDeclaration) =>
                {
                    var component = componentDeclaration;
                    var componentItem = new ExplorerItemViewModel(component);
                    foreach (var memberDeclaration in userDeclarations.Where(d => ReferenceEquals(d.ParentDeclaration, component)))
                    {
                        var member = memberDeclaration;
                        var memberItem = new ExplorerItemViewModel(member);
                        if (member.DeclarationType == DeclarationType.UserDefinedType)
                        {
                            foreach (var item in userDeclarations.Where(d => ReferenceEquals(d.ParentDeclaration, component) 
                                && d.DeclarationType == DeclarationType.UserDefinedTypeMember
                                && d.ParentScope == member.Scope))
                            {
                                memberItem.AddChild(new ExplorerItemViewModel(item));
                            }
                        }

                        if (member.DeclarationType == DeclarationType.Enumeration)
                        {
                            foreach (var item in userDeclarations.Where(d => ReferenceEquals(d.ParentDeclaration, component)
                                && d.DeclarationType == DeclarationType.EnumerationMember
                                && d.ParentScope == member.Scope))
                            {
                                memberItem.AddChild(new ExplorerItemViewModel(item));
                            }
                        }

                        componentItem.AddChild(memberItem);
                    }

                    projectItem.AddChild(componentItem);
                });

                // todo: figure out a way to yield return before that
                yield return projectItem;
            }
        }
    }
}