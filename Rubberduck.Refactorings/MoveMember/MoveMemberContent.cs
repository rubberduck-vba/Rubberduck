using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMovedMemberContentProvider
    {
        string NewPreCodeSectionContent { get; }
        string NewCodeSectionContent { get; }
        string AllNewContent { get; }
        bool HasNewContent { get; }
        int NewContentLineCount { get; }
        void AddDeclarationBlock(string declaration);
        void AddCodeBlock(string codeBlock);
        void ResetContent();
        Declaration Module { get; }
        string ModuleName { get; }
        ComponentType ComponentType { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<Declaration> ModuleDeclarations { get; }
        IMoveEndpointRewriter InsertNewContent(IMoveEndpointRewriter endpointRewriter);
        bool ContainsReference(IdentifierReference idRef);
        void AddEncapsulatedVariable(Declaration prototype, string backingVariableName, string propertyName = null);
    }

    public abstract class MoveMemberContent : IMovedMemberContentProvider
    {
        protected IProvideMoveDeclarationGroups _declarationGroups;

        public MoveMemberContent(MoveDefinition moveDefinition, IDeclarationFinderProvider declarationFinderProvider, IProvideMoveDeclarationGroups declarationGroups)
        {
            MoveDefinition = moveDefinition;
            _declarationGroups = declarationGroups;
            NewContent = new MoveMemberNewContent(moveDefinition, _declarationGroups.AllDeclarations);
        }

        protected MoveDefinition MoveDefinition { get; }

        public string NewPreCodeSectionContent => NewContent.PreCodeSectionContent;

        public string NewCodeSectionContent => NewContent.AsCodeSection;

        public string AllNewContent => NewContent.AsNewModuleContent;

        public bool HasNewContent => NewContent.HasNewContent;

        public int NewContentLineCount => NewContent.CountOfProcLines;

        protected IMovedMemberNewContentProvider NewContent { private set;  get; }

        public void AddDeclarationBlock(string declaration)
            => NewContent.AddDeclarationBlock(declaration);

        public void AddCodeBlock(string codeBlock)
            => NewContent.AddCodeBlock(codeBlock);

        public void ResetContent()
        {
            NewContent = new MoveMemberNewContent(MoveDefinition, _declarationGroups.AllDeclarations);
        }

        public void AddEncapsulatedVariable(Declaration prototype, string backingVariableName, string propertyName = null)
        {
            var propertyBlock = new PropertyBlockProvider(prototype, backingVariableName, propertyName);

            AddDeclarationBlock(propertyBlock.BackingVariableDeclaration);
            AddCodeBlock(propertyBlock.PropertyLet);
            AddCodeBlock(propertyBlock.PropertyGet);
        }

        public abstract Declaration Module { get; }
        public abstract IEnumerable<Declaration> ModuleDeclarations { get; }
        public abstract IMoveEndpointRewriter InsertNewContent(IMoveEndpointRewriter endpointRewriter);

        public virtual string ModuleName => Module.IdentifierName;
        public virtual ComponentType ComponentType => QualifiedModuleName.ComponentType;
        public QualifiedModuleName QualifiedModuleName => Module?.QualifiedModuleName ?? throw new InvalidOperationException();

        public bool ContainsReference(IdentifierReference idRef)
            => QualifiedModuleName == idRef.QualifiedModuleName;
    }

    public interface ISourceContentProvider : IMovedMemberContentProvider
    {
        bool MoveUsingExistingClassVariable { get; }
        string ClassInstantiationFragment { get; }
        string ClassInstantiationSubName { get; }
    }

    public class MoveMemberContentSource : MoveMemberContent, ISourceContentProvider
    {
        private Dictionary<Declaration, MoveMemberContentInfo> _moveMembersContentInfo = new Dictionary<Declaration, MoveMemberContentInfo>();

        public MoveMemberContentSource(MoveDefinition moveDefinition, IDeclarationFinderProvider declarationFinderProvider, IProvideMoveDeclarationGroups moveDeclarationGroups)
            : base(moveDefinition, declarationFinderProvider, moveDeclarationGroups)
        {
            ModuleDeclarations = _declarationGroups.AllDeclarations.Except(_declarationGroups.Remove);

            if (MoveDefinition.IsClassModuleDestination)
            {
                NewContent.ClassVariableContentProvider = new ClassVariableContentProvider(MoveDefinition.Destination.ModuleName);
                MoveUsingExistingClassVariable = _declarationGroups.CountOfModuleInstanceVariables(MoveDefinition.Destination.Module) == 1;
            }
        }

        public override Declaration Module => MoveDefinition.Source.Module;

        public override IEnumerable<Declaration> ModuleDeclarations { get; }

        public override IMoveEndpointRewriter InsertNewContent(IMoveEndpointRewriter endpointRewriter)
        {
            int? codeSectionStartIndex
            = ModuleDeclarations.Where(m => m.IsMember())
                        .OrderBy(c => c.Selection)
                        .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            if (codeSectionStartIndex.HasValue && HasNewContent)
            {
                if (NewPreCodeSectionContent.Length > 0)
                {
                    endpointRewriter.InsertBefore(codeSectionStartIndex.Value, NewPreCodeSectionContent);
                }

                if (NewCodeSectionContent.Length > 0)
                {
                    endpointRewriter.InsertAtEndOfFile($"{Environment.NewLine}{Environment.NewLine}{NewCodeSectionContent}");
                }
            }
            else
            {
                endpointRewriter.InsertAtEndOfFile(AllNewContent);
            }
            return endpointRewriter;
        }

        public bool MoveUsingExistingClassVariable
        {
            set
            {
                NewContent.UsesExistingClassVariable = value;
            }
            get
            {
                return NewContent.UsesExistingClassVariable;
            }
        }

        public string ClassInstantiationFragment => NewContent.ClassInstantiationFragment;

        public string ClassInstantiationSubName => NewContent.ClassInstantiationSubName;
    }

    public interface IDestinationContentProvider : IMovedMemberContentProvider
    {
        bool IsNewModule { get; }
    }

    public class MoveMemberContentDestination : MoveMemberContent, IDestinationContentProvider
    {
        public MoveMemberContentDestination(MoveDefinition moveDefinition, IDeclarationFinderProvider declarationFinderProvider, IProvideMoveDeclarationGroups declarationGroups)
            : base(moveDefinition, declarationFinderProvider, declarationGroups)
        {
            ModuleDeclarations = moveDefinition.Destination.Module == null ? Enumerable.Empty<Declaration>()
                : declarationFinderProvider.DeclarationFinder.Members(moveDefinition.Destination.Module);
        }

        public override IEnumerable<Declaration> ModuleDeclarations { get; }

        public override Declaration Module => MoveDefinition.Destination.Module;

        public bool IsNewModule => (MoveDefinition.Destination.Module is null) && MoveDefinition.Destination.ModuleName.Length > 0;

        public override string ModuleName => MoveDefinition.Destination.ModuleName;

        public override ComponentType ComponentType => MoveDefinition.Destination.ComponentType;

        public override IMoveEndpointRewriter InsertNewContent(IMoveEndpointRewriter endpointRewriter)
        {
            int? codeSectionStartIndex
                = ModuleDeclarations.Where(m => m.IsMember())
                            .OrderBy(c => c.Selection)
                            .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            endpointRewriter.InsertNewContent(codeSectionStartIndex, NewContent);
            return endpointRewriter;
        }
    }
}
