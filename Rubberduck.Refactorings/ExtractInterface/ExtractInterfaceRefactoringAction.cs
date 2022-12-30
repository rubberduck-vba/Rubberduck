using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Resources;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoringAction : RefactoringActionWithSuspension<ExtractInterfaceModel>
    {
        private readonly ICodeOnlyRefactoringAction<AddInterfaceImplementationsModel> _addImplementationsRefactoringAction;
        private readonly IParseTreeProvider _parseTreeProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IAddComponentService _addComponentService;
        private readonly IRewritingManager _rewritingManager;

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
            _rewritingManager = rewritingManager;
        }

        protected override bool RequiresSuspension(ExtractInterfaceModel model)
        {
            return true;
        }

        protected override void Refactor(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            AddInterface(model, rewriteSession, _rewritingManager.CheckOutCodePaneSession());
        }

        private void AddInterface(ExtractInterfaceModel model, IRewriteSession rewriteSession, IRewriteSession scratchPadRewriteSession)
        {
            var targetProject = _projectsProvider.Project(model.TargetDeclaration.ProjectId);
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            ModifyMembers(model, rewriteSession);
            AddInterfaceClass(model);
            AddImplementsStatement(model, rewriteSession);

            var interfaceMemberImplementations = InterfaceMemberBlocks(model, scratchPadRewriteSession);
            AddInterfaceMembersToClass(model, rewriteSession, interfaceMemberImplementations);
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
            var interfaceMembers = string.Join(NewLines.DOUBLE_SPACE, model.SelectedMembers.Select(m => m.Body));
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

            return $"{folderAnnotationText}{exposedAnnotationText}{interfaceAnnotationText}{Environment.NewLine}{optionExplicit}{Environment.NewLine}{interfaceMembers}";
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

        private void AddInterfaceMembersToClass(ExtractInterfaceModel model, IRewriteSession rewriteSession, IEnumerable<(Declaration, string)> interfaceMemberImplementations)
        {
            var addMembersModel = new AddInterfaceImplementationsModel(model.TargetDeclaration.QualifiedModuleName, model.InterfaceName, SelectedDeclarations(model));
            foreach ((Declaration member, string implementation) in interfaceMemberImplementations)
            {
                addMembersModel.SetMemberImplementation(member, implementation);
            }

            _addImplementationsRefactoringAction.Refactor(addMembersModel, rewriteSession);
        }

        private static void ModifyMembers(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            if (model.ImplementationOption == ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface)
            {
                RedirectMemberReferencesToImplementingMember(model, rewriteSession);
                ForwardMembers(model, rewriteSession, SelectedDeclarations(model));
            }
            else if (model.ImplementationOption == ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)
            {
                RedirectMemberReferencesToImplementingMember(model, rewriteSession);

                //Only delete members free of external resources - otherwise, this option would generate uncompilable code
                var membersToRetain = MembersWithExternalReferences(model);

                ForwardMembers(model, rewriteSession, membersToRetain);

                DeleteMembers(model, rewriteSession, SelectedDeclarations(model).Except(membersToRetain));
            }
        }

        private static IEnumerable<(Declaration, string)> InterfaceMemberBlocks(ExtractInterfaceModel model, IRewriteSession scratchPadRewriteSession)
        {
            switch (model.ImplementationOption)
            {
                case ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface:
                case ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface:
                    return ReplicateBlockToImplementingMember(model, scratchPadRewriteSession);
                case ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers:
                    return CreateForwardingBlocks(model);
                default:
                    return Enumerable.Empty<(Declaration, string)>();
            }
        }

        private static IEnumerable<(Declaration, string)> CreateForwardingBlocks(ExtractInterfaceModel model)
        {
            string ForwardingBlock(ModuleBodyElementDeclaration member) =>
                member.DeclarationType.HasFlag(DeclarationType.Function)
                    ? ForwardFunction(member.IdentifierName, model.ImplementingMemberName(member.IdentifierName), member)
                    : ForwardProcedure(member.IdentifierName, member);

            return SelectedDeclarations(model).OfType<ModuleBodyElementDeclaration>()
                .Select(m => (m as Declaration, ForwardingBlock(m)));
        }

        private static IEnumerable<(Declaration, string)> ReplicateBlockToImplementingMember(ExtractInterfaceModel model, IRewriteSession scratchPadRewriteSession)
        {
            var selectedDeclarations = SelectedDeclarations(model);
            if (!selectedDeclarations.Any())
            {
                return Enumerable.Empty<(Declaration, string)>();
            }
            var scratchPadRewriter = scratchPadRewriteSession.CheckOutModuleRewriter(selectedDeclarations.First().QualifiedModuleName);
            foreach (IdentifierReference identifierReference in selectedDeclarations.SelectMany(m => m.References).Where(rf => rf.QualifiedModuleName == rf.Declaration.QualifiedModuleName))
            {
                scratchPadRewriter.Replace(identifierReference.Context, model.ImplementingMemberName(identifierReference.IdentifierName));
            }

            var implementations =  selectedDeclarations.OfType<ModuleBodyElementDeclaration>()
                .Where(m => m.Block.ContainsExecutableStatements(true))
                .Select(m => (m as Declaration, $"{scratchPadRewriter.GetText(m.Block.Start.TokenIndex, m.Block.Stop.TokenIndex).Trim()}"))
                .ToList();

            if (model.ImplementationOption == ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface
                && selectedDeclarations.OfType<ModuleBodyElementDeclaration>().Count() > implementations.Count())
            {
                var membersToRetain = MembersWithExternalReferences(model);
                foreach (var toRetain in membersToRetain)
                {
                    if (!implementations.Select(imp => imp.Item1).Contains(toRetain))
                    {
                        implementations.Add((toRetain, Resources.Refactorings.Refactorings.ImplementInterface_TODO));
                    }
                }
            }

            return implementations;
        }

        private static void RedirectMemberReferencesToImplementingMember(ExtractInterfaceModel model, IRewriteSession rewriteSession)
        {
            var selectedDeclarations = SelectedDeclarations(model);
            var nonSelectedMembers = model.DeclarationFinderProvider.DeclarationFinder.Members(model.TargetDeclaration.QualifiedModuleName)
                .OfType<ModuleBodyElementDeclaration>()
                .Except(selectedDeclarations);

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);
            foreach (var member in nonSelectedMembers)
            {
                var otherInterfaceMemberReferences = selectedDeclarations.SelectMany(m => m.References).Where(rf => rf.ParentScoping == member);
                foreach (IdentifierReference identifierReference in otherInterfaceMemberReferences)
                {
                    rewriter.Replace(identifierReference.Context, model.ImplementingMemberName(identifierReference.IdentifierName));
                }
            }
        }

        private static void ForwardMembers(ExtractInterfaceModel model, IRewriteSession rewriteSession, IEnumerable<Declaration> membersToForward)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            foreach (var member in membersToForward.OfType<ModuleBodyElementDeclaration>())
            {
                var forwardingStatement = member.DeclarationType.HasFlag(DeclarationType.Function)
                    ? ForwardFunction(model.ImplementingMemberName(member.IdentifierName), member.IdentifierName, member)
                    : ForwardProcedure(model.ImplementingMemberName(member.IdentifierName), member);

                if (member.Block.ContainsExecutableStatements(true))
                {
                    rewriter.Replace(member.Block, $"{forwardingStatement}{Environment.NewLine}");
                    continue;
                }
                rewriter.InsertBefore(member.Block.Start.TokenIndex, $"    {forwardingStatement}{Environment.NewLine}");
            }
        }

        private static string ForwardFunction(string targetMemberIdentifier, string forwardingMemberIdentifier, ModuleBodyElementDeclaration member)
        {
                var forwardStatementLHS = !member.IsObject
                ? forwardingMemberIdentifier
                : $"{Tokens.Set} {forwardingMemberIdentifier}";

            var forwardStatementRHS = IsParameterlessPropertyGet(member)
                ? targetMemberIdentifier
                : $"{targetMemberIdentifier}({string.Join(", ", member.Parameters.Select(p => p.IdentifierName))})";

            return $"{forwardStatementLHS} = {forwardStatementRHS}";
        }

        private static string ForwardProcedure(string targetMemberIdentifier, ModuleBodyElementDeclaration member)
        {
            if (member.DeclarationType.Equals(DeclarationType.Procedure))
            {
                return ($"{targetMemberIdentifier} {string.Join(", ", member.Parameters.Select(p => p.IdentifierName))}");
            }

            var forwardStatementLHS = (member.Parameters.Count != 1)
                                            ? $"{targetMemberIdentifier}{$"({string.Join(", ", member.Parameters.Take(member.Parameters.Count - 1).Select(p => p.IdentifierName))})"}"
                                            : targetMemberIdentifier;

            var forwardStatementRHS = $"{member.Parameters.Last().IdentifierName}";

            return member.DeclarationType.Equals(DeclarationType.PropertySet)
                ? $"{Tokens.Set} {forwardStatementLHS} = {forwardStatementRHS}"
                : $"{forwardStatementLHS} = {forwardStatementRHS}";
        }

        private static void DeleteMembers(ExtractInterfaceModel model, IRewriteSession rewriteSession, IEnumerable<Declaration> membersToDelete)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            foreach (var member in membersToDelete)
            {
                var moduleBodyElementContext = member.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
                rewriter.Remove(moduleBodyElementContext);
            }

            var moduleBodyContext = model.SelectedMembers.FirstOrDefault()?.Member.Context.GetAncestor<VBAParser.ModuleBodyContext>();
            if (moduleBodyContext != null)
            {
                var moduleBodyContent = rewriter.GetText(moduleBodyContext.Start.TokenIndex, moduleBodyContext.Stop.TokenIndex);
                moduleBodyContent = ConstrainNewlineSequences(moduleBodyContent, 2).Trim();
                Interval moduleBodyInterval = new Interval(moduleBodyContext.Start.TokenIndex, moduleBodyContext.Stop.TokenIndex);
                rewriter.Replace(moduleBodyInterval, moduleBodyContent);
            }
        }

        private static bool IsParameterlessPropertyGet(ModuleBodyElementDeclaration member)
            => member.DeclarationType.Equals(DeclarationType.PropertyGet) && member.Parameters.Count == 0;

        private static List<Declaration> SelectedDeclarations(ExtractInterfaceModel model)
            => model.SelectedMembers.Select(m => m.Member).ToList();

        private static List<Declaration> MembersWithExternalReferences(ExtractInterfaceModel model)
            => SelectedDeclarations(model).Where(d => d.References.Any(rf => rf.QualifiedModuleName != d.QualifiedModuleName)).ToList();


        private static string ConstrainNewlineSequences(string content, int maxConsecutiveNewLines)
        {
            if (maxConsecutiveNewLines <= 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            var targetNewlineSequence = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines + 1));
            var maxNewlineSequence = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines));
            var guard = 0;
            while (content.Contains(targetNewlineSequence) && ++guard < 100)
            {
                content = content.Replace(targetNewlineSequence, maxNewlineSequence);
            }
            return content;
        }
    }
}
