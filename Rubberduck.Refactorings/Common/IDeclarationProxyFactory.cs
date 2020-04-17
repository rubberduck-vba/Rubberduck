using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationProxy
    {
        string IdentifierName { set; get; }
        string TargetModuleName { set; get; }
        Declaration TargetModule { set; get; }
        Declaration Prototype { get; }
        DeclarationType DeclarationType { set; get; }
        IEnumerable<IdentifierReference> References { get; }
        Declaration ParentDeclaration { set; get; }
        Accessibility Accessibility { set; get; }
        string ProjectId { get; }
        string ProjectName { get; }
    }

    public interface IDeclarationProxyFactory
    {
        IDeclarationProxy Create(Declaration prototype, string identifierName, ModuleDeclaration targetModule);
        IDeclarationProxy Create(string identifier, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration targetModule, Declaration parentDeclaration);
    }

    public class DeclarationProxyFactory : IDeclarationProxyFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public DeclarationProxyFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IDeclarationProxy Create(string identifier, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration targetModule, Declaration parentDeclaration)
        {
            return new DeclarationProxy(identifier, declarationType, accessibility, targetModule, parentDeclaration);
        }

        public IDeclarationProxy Create(Declaration prototype, string identifierName, ModuleDeclaration targetModule)
        {
            return new DeclarationProxy(prototype, identifierName, targetModule);
        }
    }
}
