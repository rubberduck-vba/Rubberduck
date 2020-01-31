using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveScenario
    {
        ISourceContentProvider SourceContentProvider { get; }
        IDestinationContentProvider DestinationContentProvider { get; }
        MoveDefinition MoveDefinition { set; get; }
        IDeclarationFinderProvider DeclarationFinderProvider { get; }
        IEnumerable<Declaration> SelectedDeclarations { get; }
        MoveElementGroups SelectedElements { get; }
        bool IsValidMoveDefinition { get; }
        bool CreatesNewModule { get; }
        MoveEndpoints Endpoints { get; }
        bool ForwardSelectedMemberCalls { set; get; }
        bool IsMoveEndpoint(QualifiedModuleName qmn);
        bool IsAsSelectedVariable(Declaration element);
        bool IsAsSelectedMember(Declaration element);
        bool IsOnlyReferencedByMovedElements(Declaration element);
        QualifiedModuleName QualifiedModuleNameSource { get; }
    }

    public class MoveScenario : IMoveScenario, IProvideMoveDeclarationGroups
    {
        public MoveDefinition MoveDefinition { set; get; }

        public MoveScenario(MoveDefinition moveDefinition, IDeclarationFinderProvider declarationFinderProvider)
        {
            DeclarationFinderProvider = declarationFinderProvider;

            MoveDefinition = moveDefinition;

            SourceDeclarationGroups = new MoveMemberDeclarationGroups(moveDefinition, declarationFinderProvider);

            SourceContentProvider = new MoveMemberContentSource(moveDefinition, declarationFinderProvider, SourceDeclarationGroups);

            DestinationContentProvider = new MoveMemberContentDestination(moveDefinition, declarationFinderProvider, SourceDeclarationGroups);
        }

        public IDeclarationFinderProvider DeclarationFinderProvider { get; }

        public ISourceContentProvider SourceContentProvider { set;  get; }
        public IDestinationContentProvider DestinationContentProvider { set;  get; }
        private IProvideMoveDeclarationGroups SourceDeclarationGroups { set; get; }

        public QualifiedModuleName QualifiedModuleNameSource
            => MoveDefinition.Source.QualifiedModuleName.Value;

        public bool ForwardSelectedMemberCalls
        {
            set
            {
                SourceDeclarationGroups.ForwardSelectedMemberCalls = value;
            }
            get
            {
                return SourceDeclarationGroups.ForwardSelectedMemberCalls;
            }
        }

        public IEnumerable<Declaration> SelectedDeclarations
            => SelectedElements.AllDeclarations;

        public bool IsAsSelectedVariable(Declaration element)
            => SelectedElements.Contains(element) && element.IsVariable();

        public bool IsAsSelectedMember(Declaration element)
            => SelectedElements.Contains(element) && element.IsMember();

        public bool IsOnlyReferencedByMovedElements(Declaration element)
            => element.References.All(rf => Moving.Members.Contains(rf.ParentScoping));

        public bool IsMoveEndpoint(QualifiedModuleName qmn)
        {
            return qmn == SourceContentProvider.QualifiedModuleName || qmn == DestinationContentProvider.QualifiedModuleName;
        }

        public bool IsPropertyBackingVariable(Declaration variable) => SourceDeclarationGroups.IsPropertyBackingVariable(variable);

        public bool IsSingleDeclarationSelection => SourceDeclarationGroups.IsSingleDeclarationSelection;

        public bool TryGetPropertiesFromBackingVariable(Declaration variable, out List<Declaration> properties)
        {
            return SourceDeclarationGroups.TryGetPropertiesFromBackingVariable(variable, out properties);
        }

        public MoveEndpoints Endpoints => MoveDefinition.Endpoints;

        protected IEnumerable<Declaration> DestinationClassInstanceVariables
        { 
            get
            {
                if (CreatesNewModule) { return Enumerable.Empty<Declaration>(); }

                return AllDeclarations.Where(el => el.IsVariable()
                    && (el.AsTypeDeclaration?.IdentifierName.Equals(DestinationContentProvider.Module.IdentifierName) ?? false));
            }
        }

        public static IMoveScenario NullMove() => new NullMoveScenario();

        public bool IsValidMoveDefinition
        {
            get
            {
                return !(DestinationEqualsSourceModuleName()
                    || HasNoDestinationSpecified()
                    //|| CausesDestinationNameConflicts(DeclarationFinderProvider, this as IProvideMoveDeclarationGroups)
                    || IsInvalidDestinationModuleName(MoveDefinition.Destination.ModuleName));
            }
        }

        public bool CreatesNewModule => DestinationContentProvider.IsNewModule;

        public MoveElementGroups Moving => SourceDeclarationGroups.Moving;

        public MoveElementGroups SelectedElements => SourceDeclarationGroups.SelectedElements;

        public IEnumerable<Declaration> MoveableElements { get => SourceDeclarationGroups.MoveableElements; set => SourceDeclarationGroups.MoveableElements = value; }

        public MoveElementGroups Participants => SourceDeclarationGroups.Participants;

        public MoveElementGroups SupportingElements => SourceDeclarationGroups.SupportingElements;

        public MoveElementGroups Retain => SourceDeclarationGroups.Retain;

        public Dictionary<Declaration, string> VariableReferenceReplacement { get => SourceDeclarationGroups.VariableReferenceReplacement; set => SourceDeclarationGroups.VariableReferenceReplacement = value; }

        public MoveElementGroups MoveAndDelete => SourceDeclarationGroups.MoveAndDelete;

        public IEnumerable<Declaration> Forward => SourceDeclarationGroups.Forward;

        public IEnumerable<Declaration> Remove => SourceDeclarationGroups.Remove;

        public IEnumerable<Declaration> AllDeclarations => SourceDeclarationGroups.AllDeclarations;

        public int CountOfModuleInstanceVariables(Declaration module) => SourceDeclarationGroups.CountOfModuleInstanceVariables(module);

        public bool InternalStateUsedByMovedMembers => SourceDeclarationGroups.InternalStateUsedByMovedMembers;

        //public static bool DeclarationCanBeAnalyzed(Declaration declaration)
        //{
        //    return (MoveableTypeFlags.Any(flag => declaration.DeclarationType.HasFlag(flag))
        //        || declaration.DeclarationType.HasFlag(DeclarationType.Event))
        //        && !declaration.IsLocalConstant()
        //        && !declaration.IsLocalVariable();
        //}

        private bool DestinationEqualsSourceModuleName()
            => MoveDefinition.Source.ModuleName.Equals(MoveDefinition.Destination.ModuleName);

        private bool HasNoDestinationSpecified()
            => string.IsNullOrEmpty(MoveDefinition.Destination.ModuleName);

        private bool CausesDestinationNameConflicts(IDeclarationFinderProvider declarationFinderProvider, IProvideMoveDeclarationGroups groups)
        {
            if (MoveDefinition.Destination.Module is null) { return false; }

            var destinationDeclarations = MoveDefinition.Destination.Module != null ?
                declarationFinderProvider.DeclarationFinder.Members(MoveDefinition.Destination.Module)
                : Enumerable.Empty<Declaration>();

            var details = destinationDeclarations.Where(dec => groups.MoveAndDelete.AllDeclarations.Any(nm => nm.IdentifierName.Equals(dec.IdentifierName)));
            return details.Any();
        }

        //private static List<(Func<string, bool>, string)> ChecksForInvalidNames = new List<(Func<string, bool>, string)>()
        //{
        //    (ModuleNameValidator.StartsWithDigit, "'{0}' does not start with a letter"),
        //    (ModuleNameValidator.IsReservedIdentifier, "'{0}' is a VBA reserved identifier"),
        //    (ModuleNameValidator.HasSpecialCharacters, "'{0}' contains special characters"),
        //    (ModuleNameValidator.IsOverMaxLength, $"'{"{0}'"} exceeds {ModuleNameValidator.MaxNameLength} characters"),
        //};

        private bool IsInvalidDestinationModuleName(string destinationModuleName)
        {
            if (string.IsNullOrEmpty(MoveDefinition.Destination.ModuleName))
            {
                return false;
            }

            return VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(destinationModuleName, DeclarationType.Module, out var criteriaMatchMessage);
            //var descriptors = ChecksForInvalidNames.Where(s => s.Item1(destinationModuleName));
            //return descriptors.Any();
        }

        //TODO: Review the two interfaces for overlapping members....required in both interfaces?
        private class NullMoveScenario : IMoveScenario, IProvideMoveDeclarationGroups
        {
            public ISourceContentProvider SourceContentProvider { get; }
            public IDestinationContentProvider DestinationContentProvider { get; }
            public MoveDefinition MoveDefinition { set; get; }
            private IProvideMoveDeclarationGroups SourceDeclarationGroups { get; }
            public MoveElementGroups SelectedElements { get; } = new MoveElementGroups(Enumerable.Empty<Declaration>());
            public IDeclarationFinderProvider DeclarationFinderProvider { get; }
            public bool IsValidMoveDefinition => false;
            public IEnumerable<Declaration> SelectedDeclarations => Enumerable.Empty<Declaration>();
            public bool CreatesNewModule => false;
            void ModifySourceModule() { }
            public MoveEndpoints Endpoints => MoveEndpoints.Undefined;
            public bool ForwardSelectedMemberCalls { set; get; }
            public bool IsMoveEndpoint(QualifiedModuleName qmn) => false;
            public bool IsAsSelectedVariable(Declaration element) => false;
            public bool IsAsSelectedMember(Declaration element) => false;
            public bool IsOnlyReferencedByMovedElements(Declaration element) => false;
            public bool ForwardCall => false;
            public QualifiedModuleName QualifiedModuleNameSource => new QualifiedModuleName();

            public MoveElementGroups Moving { get; } = new MoveElementGroups(Enumerable.Empty<Declaration>());
            public IEnumerable<Declaration> MoveableElements { set; get; } = Enumerable.Empty<Declaration>();
            public MoveElementGroups Participants { get; } = new MoveElementGroups(Enumerable.Empty<Declaration>());
            public MoveElementGroups SupportingElements { get; } = new MoveElementGroups(Enumerable.Empty<Declaration>());
            public MoveElementGroups Retain { get; } = new MoveElementGroups(Enumerable.Empty<Declaration>());
            public Dictionary<Declaration, string> VariableReferenceReplacement { set; get; }
            public MoveElementGroups MoveAndDelete { get; } = new MoveElementGroups(Enumerable.Empty<Declaration>());
            public IEnumerable<Declaration> Forward { get; } = Enumerable.Empty<Declaration>();
            public IEnumerable<Declaration> Remove { get; } = Enumerable.Empty<Declaration>();
            public IEnumerable<Declaration> AllDeclarations { get; } = Enumerable.Empty<Declaration>();
            public bool IsPropertyBackingVariable(Declaration variable) => false;
            public bool TryGetPropertiesFromBackingVariable(Declaration variable, out List<Declaration> properties){ properties = null; return false;}
            public bool InternalStateUsedByMovedMembers { get; } = false;
            public int CountOfModuleInstanceVariables(Declaration module) => 0;
            public bool IsSingleDeclarationSelection { get; } = false;
        }
    }
}
