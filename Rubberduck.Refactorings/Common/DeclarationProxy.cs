using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public class DeclarationProxy : IDeclarationProxy
    {
        private readonly Declaration _declaration;

        public DeclarationProxy(Declaration prototype, string identifier, ModuleDeclaration targetModule)
            : this(identifier, prototype.DeclarationType, prototype.Accessibility, targetModule, targetModule)
        {
            _declaration = prototype;
            TargetModule = targetModule;
            TargetModuleName = TargetModule.IdentifierName;
            ParentDeclaration = _declaration.ParentDeclaration is ModuleDeclaration
                                        ? TargetModule
                                        : _declaration.ParentDeclaration;
        }

        public DeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration targetModule, Declaration parentDeclaration)
        {
            TargetModule = targetModule;
            IdentifierName = identifier;
            DeclarationType = declarationType;
            Accessibility = accessibility;
            ParentDeclaration = parentDeclaration;
        }

        public Declaration Prototype => _declaration;

        public Declaration TargetModule { set; get; }

        public Declaration ParentDeclaration { set; get; }

        public string IdentifierName { set; get; }

        public DeclarationType DeclarationType { set; get; }

        public IEnumerable<IdentifierReference> References 
            => _declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        public string ProjectId => TargetModule.ProjectId ?? string.Empty;

        public string ProjectName => TargetModule.ProjectName ?? string.Empty;

        public Accessibility Accessibility { set; get; }

        public string TargetModuleName { set; get; }
    }
}
