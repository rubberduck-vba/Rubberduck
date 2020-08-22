using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoringAction : RefactoringActionWithSuspension<ExtractInterfaceModel>
    {
        private readonly ICodeOnlyRefactoringAction<AddInterfaceImplementationsModel> _addImplementationsRefactoringAction;
        private readonly IParseTreeProvider _parseTreeProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IAddComponentService _addComponentService;

        public ExtractInterfaceRefactoringAction(
            AddInterfaceImplementationsRefactoringAction addImplementationsRefactoringAction,
            IParseTreeProvider parseTreeProvider,
            IParseManager parseManager, 
            IRewritingManager rewritingManager,
            IProjectsProvider projectsProvider,
            IAddComponentService addComponentService) 
            : base(parseManager, rewritingManager)
        {
            _addImplementationsRefactoringAction = addImplementationsRefactoringAction;
            _parseTreeProvider = parseTreeProvider;
            _projectsProvider = projectsProvider;
            _addComponentService = addComponentService;
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
            var targetProject = _projectsProvider.Project(model.TargetDeclaration.ProjectId);
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            AddInterfaceClass(model);
            AddImplementsStatement(model, rewriteSession);
            AddInterfaceMembersToClass(model, rewriteSession);
        }

        private void AddInterfaceClass(ExtractInterfaceModel model)
        {
            var targetProjectId = model.TargetDeclaration.ProjectId;
            var interfaceCode = InterfaceCode(model);
            var interfaceName = model.InterfaceName;

            if (model.InterfaceInstancing == ClassInstancing.Public)
            {
                _addComponentService.AddComponentWithAttributes(targetProjectId, ComponentType.ClassModule, interfaceCode, componentName: interfaceName);
            }
            else
            {
                _addComponentService.AddComponent(targetProjectId, ComponentType.ClassModule, interfaceCode, componentName: interfaceName);
            }
        }

        private static string InterfaceCode(ExtractInterfaceModel model)
        {
            var interfaceBody = InterfaceModuleBody(model);

            if (model.InterfaceInstancing == ClassInstancing.Public)
            {
                var moduleHeader = ExposedInterfaceHeader(model.InterfaceName);
                return $"{moduleHeader}{Environment.NewLine}{interfaceBody}";
            }

            return interfaceBody;
        }

        private static string InterfaceModuleBody(ExtractInterfaceModel model)
        {
            var interfaceMembers = string.Join(Environment.NewLine, model.SelectedMembers.Select(m => m.Body));
            var optionExplicit = $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}";

            var targetModule = Declaration.GetModuleParent(model.TargetDeclaration);
            var folderAnnotation = targetModule?.Annotations.FirstOrDefault(pta => pta.Annotation is FolderAnnotation);
            var folderAnnotationText = folderAnnotation != null
                                       ? $"'@{folderAnnotation.Context.GetText()}{Environment.NewLine}"
                                       : string.Empty;

            var exposedAnnotation = new ExposedModuleAnnotation();
            var exposedAnnotationText = model.InterfaceInstancing == ClassInstancing.Public
                ? $"'@{exposedAnnotation.Name}{Environment.NewLine}"
                : string.Empty;

            var interfaceAnnotation = new InterfaceAnnotation();
            var interfaceAnnotationText = $"'@{interfaceAnnotation.Name}{Environment.NewLine}";

            return $"{optionExplicit}{Environment.NewLine}{folderAnnotationText}{exposedAnnotationText}{interfaceAnnotationText}{Environment.NewLine}{interfaceMembers}";
        }

        private static string ExposedInterfaceHeader(string interfaceName)
        {
            return $@"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""{interfaceName}""
Attribute VB_Exposed = True";
        }

        private void AddImplementsStatement(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            var implementsStatement = $"Implements {model.InterfaceName}";

            var (insertionIndex, isImplementsStatement) = InsertionIndex(model);

            if (insertionIndex == -1)
            {
                rewriter.InsertBefore(0, $"{implementsStatement}{Environment.NewLine}{Environment.NewLine}");
            }
            else
            {
                rewriter.InsertAfter(insertionIndex, $"{Environment.NewLine}{(isImplementsStatement ? string.Empty : Environment.NewLine)}{implementsStatement}");
            }
        }

        private (int index, bool isImplementsStatement) InsertionIndex(ExtractInterfaceModel model)
        {
            var tree = (ParserRuleContext)_parseTreeProvider.GetParseTree(model.TargetDeclaration.QualifiedModuleName, CodeKind.CodePaneCode);

            var moduleDeclarations = tree.GetDescendent<VBAParser.ModuleDeclarationsContext>();
            if (moduleDeclarations == null)
            {
                return (-1, false);
            }

            var lastImplementsStatement = moduleDeclarations
                .GetDescendents<VBAParser.ImplementsStmtContext>()
                .LastOrDefault();
            if (lastImplementsStatement != null)
            {
                return (lastImplementsStatement.Stop.TokenIndex, true);
            }

            var lastOptionStatement = moduleDeclarations
                .GetDescendents<VBAParser.ModuleOptionContext>()
                .LastOrDefault();
            if (lastOptionStatement != null)
            {
                return (lastOptionStatement.Stop.TokenIndex, false);
            }

            return (-1, false);
        }

        private void AddInterfaceMembersToClass(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            var targetModule = model.TargetDeclaration.QualifiedModuleName;
            var interfaceName = model.InterfaceName;
            var membersToImplement = model.SelectedMembers.Select(m => m.Member).ToList();

            var addMembersModel = new AddInterfaceImplementationsModel(targetModule, interfaceName, membersToImplement);
            _addImplementationsRefactoringAction.Refactor(addMembersModel, rewriteSession);
        }
    }
}