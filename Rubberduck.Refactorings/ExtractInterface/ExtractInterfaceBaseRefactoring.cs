using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceBaseRefactoring : BaseRefactoringWithSuspensionBase<ExtractInterfaceModel>
    {
        private readonly ICodeOnlyBaseRefactoring<AddInterfaceImplementationsModel> _addImplementationsRefactoring;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ExtractInterfaceBaseRefactoring(
            AddInterFaceImplementationsBaseRefactoring addImplementationsRefactoring,
            IDeclarationFinderProvider declarationFinderProvider,
            IParseManager parseManager, 
            IRewritingManager rewritingManager) 
            : base(parseManager, rewritingManager)
        {
            _addImplementationsRefactoring = addImplementationsRefactoring;
            _declarationFinderProvider = declarationFinderProvider;
        }

        protected override bool RequiresSuspension(ExtractInterfaceModel model)
        {
            return true;
        }

        protected override void Refactor(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            AddInterface(model, rewriteSession);
        }

        private void AddInterface(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            var targetProject = model.TargetDeclaration.Project;
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            AddInterfaceClass(model.TargetDeclaration, model.InterfaceName, GetInterfaceModuleBody(model));

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            var firstNonFieldMember = _declarationFinderProvider.DeclarationFinder.Members(model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

            AddInterfaceMembersToClass(model, rewriteSession);
        }

        private void AddInterfaceClass(Declaration implementingClass, string interfaceName, string interfaceBody)
        {
            var targetProject = implementingClass.Project;
            using (var components = targetProject.VBComponents)
            {
                using (var interfaceComponent = components.Add(ComponentType.ClassModule))
                {
                    using (var interfaceModule = interfaceComponent.CodeModule)
                    {
                        interfaceComponent.Name = interfaceName;

                        var optionPresent = interfaceModule.CountOfLines > 1;
                        if (!optionPresent)
                        {
                            interfaceModule.InsertLines(1, $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}");
                        }
                        interfaceModule.InsertLines(3, interfaceBody);
                    }
                }
            }
        }

        private void AddInterfaceMembersToClass(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            var targetModule = model.TargetDeclaration.QualifiedModuleName;
            var interfaceName = model.InterfaceName;
            var membersToImplement = model.SelectedMembers.Select(m => m.Member).ToList();

            var addMembersModel = new AddInterfaceImplementationsModel(targetModule, interfaceName, membersToImplement);
            _addImplementationsRefactoring.Refactor(addMembersModel, rewriteSession);
        }

        private static string GetInterfaceModuleBody(ExtractInterfaceModel model)
        {
            return string.Join(Environment.NewLine, model.SelectedMembers.Select(m => m.Body));
        }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.ClassModule,
            DeclarationType.Document,
            DeclarationType.UserForm
        };
    }
}