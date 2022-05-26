using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameComponentOrProjectRefactoringAction : RefactoringActionWithSuspension<RenameModel>
    {
        private const string AppendUnderscoreFormat = "{0}_";

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IProjectsProvider _projectsProvider;

        public RenameComponentOrProjectRefactoringAction(
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider,
            IParseManager parserManager,
            IRewritingManager rewritingManager)
            : base(parserManager, rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _projectsProvider = projectsProvider;
        }

        protected override bool RequiresSuspension(RenameModel model)
        {
            //The parser needs to be suspended during the refactoring of a component because the VBE API object rename causes a separate reparse.
            return true;
        }

        protected override void Refactor(RenameModel model, IRewriteSession rewriteSession)
        {
            var targetDeclarationType = model.Target.DeclarationType;
            if (targetDeclarationType.HasFlag(DeclarationType.Module))
            {
                RenameModule(model, rewriteSession);
            }
            else if (targetDeclarationType.HasFlag(DeclarationType.Project))
            {
                RenameProject(model, rewriteSession);
            }
        }

        private void RenameModule(RenameModel model, IRewriteSession rewriteSession)
        {
            RenameReferences(model.Target, model.NewName, rewriteSession);

            if (model.Target.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                foreach (var reference in model.Target.References)
                {
                    var ctxt = reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>();
                    if (ctxt != null)
                    {
                        RenameDefinedFormatMembers(model, _declarationFinderProvider.DeclarationFinder.FindInterfaceMembersForImplementsContext(ctxt).ToList(), AppendUnderscoreFormat, rewriteSession);
                    }
                }
            }

            var component = _projectsProvider.Component(model.Target.QualifiedName.QualifiedModuleName);
            switch (component.Type)
            {
                case ComponentType.Document:
                    {
                        using (var properties = component.Properties)
                        using (var property = properties["_CodeName"])
                        {
                            property.Value = model.NewName;
                        }
                        break;
                    }
                case ComponentType.UserForm:
                case ComponentType.VBForm:
                case ComponentType.MDIForm:
                    {
                        using (var properties = component.Properties)
                        using (var property = properties["Caption"])
                        {
                            if ((string)property.Value == model.Target.IdentifierName)
                            {
                                property.Value = model.NewName;
                            }
                            component.Name = model.NewName;
                        }
                        break;
                    }
                default:
                    {
                        using (var vbe = component.VBE)
                        {
                            if (vbe.Kind == VBEKind.Hosted)
                            {
                                // VBA - rename code module
                                using (var codeModule = component.CodeModule)
                                {
                                    Debug.Assert(!codeModule.IsWrappingNullReference,
                                        "input validation fail: Attempting to rename an ICodeModule wrapping a null reference");
                                    codeModule.Name = model.NewName;
                                }
                            }
                            else
                            {
                                // VB6 - rename component
                                component.Name = model.NewName;
                            }
                        }
                        break;
                    }
            }
        }

        private void RenameProject(RenameModel model, IRewriteSession rewriteSession)
        {
            var project = _projectsProvider.Project(model.Target.ProjectId);

            if (project != null)
            {
                project.Name = model.NewName;
            }
            RenameReferences(model.Target, model.NewName, rewriteSession);
        }

        private void RenameDefinedFormatMembers(RenameModel model, IReadOnlyCollection<Declaration> members, string underscoreFormat, IRewriteSession rewriteSession)
        {
            if (!members.Any()) { return; }

            var targetFragment = string.Format(underscoreFormat, model.Target.IdentifierName);
            var replacementFragment = string.Format(underscoreFormat, model.NewName);
            foreach (var member in members)
            {
                var newMemberName = member.IdentifierName.Replace(targetFragment, replacementFragment);
                RenameStandardElements(member, newMemberName, rewriteSession);
            }
        }

        private void RenameStandardElements(Declaration target, string newName, IRewriteSession rewriteSession)
        {
            RenameReferences(target, newName, rewriteSession);
            RenameDeclaration(target, newName, rewriteSession);
        }

        private void RenameReferences(Declaration target, string newName, IRewriteSession rewriteSession)
        {
            var modules = target.References
                .Where(reference =>
                    reference.Context.GetText() != "Me"
                    && !reference.IsArrayAccess
                    && !reference.IsDefaultMemberAccess)
                .GroupBy(r => r.QualifiedModuleName);

            foreach (var grouping in modules)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(grouping.Key);
                foreach (var reference in grouping)
                {
                    rewriter.Replace(reference.Context, newName);
                }
            }
        }

        private void RenameDeclaration(Declaration target, string newName, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedName.QualifiedModuleName);

            if (target.Context is IIdentifierContext context)
            {
                rewriter.Replace(context.IdentifierTokens, newName);
            }
        }
    }
}