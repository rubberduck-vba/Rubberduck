using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Resources;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace Rubberduck.UI.Refactorings.MoveMember
{

    public class MoveMemberViewModel : RefactoringViewModelBase<MoveMemberModel>
    {
        public enum PreviewModule { Source, Destination }

        private List<MoveableMemberSetViewModel> _moveableMemberSetViewModels;
        private List<string> _existingModuleNames;

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMoveMemberRefactoringPreviewerFactory _previewerFactory;
        public MoveMemberViewModel(MoveMemberModel model, 
                                    IDeclarationFinderProvider declarationFinderProvider,
                                    IMoveMemberRefactoringPreviewerFactory previewerFactory)
            : base(model)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _previewerFactory = previewerFactory;
            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SetAllSelections(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SetAllSelections(false));

            _moveableMemberSetViewModels = new List<MoveableMemberSetViewModel>();
            foreach (var moveableMemberSet in model.MoveableMemberSets)
            {
                if ((moveableMemberSet.Member.IsVariable() && moveableMemberSet.Member.HasPrivateAccessibility())
                    || moveableMemberSet.IsUserDefinedType
                    || moveableMemberSet.IsEnumeration
                    || moveableMemberSet.Member.IsLifeCycleHandler())
                {
                    continue;
                }
                var moveableMemberViewModel = new MoveableMemberSetViewModel(RefreshPreview, moveableMemberSet);
                _moveableMemberSetViewModels.Add(moveableMemberViewModel);
            }
            _moveableMembers = new ObservableCollection<MoveableMemberSetViewModel>(OrderedMemberSets());

            _existingModuleNames = _declarationFinderProvider.DeclarationFinder.AllUserDeclarations
                .OfType<ModuleDeclaration>()
                .Select(m => m.IdentifierName).ToList();

            _previewSelection = PreviewModule.Destination;
            _destinationNameFailureCriteria = string.Empty;
        }

        public bool IsExecutableMove
        {
            get
            {
                var result = IsValidModuleName && Model.IsExecutable;
                OnPropertyChanged(nameof(IsValidModuleName));
                OnPropertyChanged(nameof(DestinationNameFailureCriteria));
                return result;
            }
        }

        private string _destinationNameFailureCriteria;
        public string DestinationNameFailureCriteria
        {
            get
            {
                return _destinationNameFailureCriteria ?? string.Empty;
            }
        }

        public bool IsValidModuleName
        {
            get
            {
                var isValid = TryValidateDestinationModuleName(out _destinationNameFailureCriteria);
                OnPropertyChanged(nameof(DestinationNameFailureCriteria));
                return isValid;
            }
        }

        private bool TryValidateDestinationModuleName(out string failCriteria)
        {
            failCriteria = string.Empty;
            if (VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(DestinationModuleName, DeclarationType.ProceduralModule, out var criteriaMessage))
            {
                failCriteria = criteriaMessage;
                return false;
            }

            if (AreVBAEquivalent(Model.Source.Module.ProjectName, DestinationModuleName))
            {
                failCriteria = RubberduckUI.MoveMember_ModuleMatchesProjectNameFailMsg;
                return false;
            }

            if (AreVBAEquivalent(Model.Source.ModuleName, DestinationModuleName))
            {
                failCriteria = RubberduckUI.MoveMember_SourceAndDestinationModuleNameMatchFailMsg;
                return false;
            }
            return true;
        }

        public string SourceModuleName => $"{Model.Source.ModuleName}";

        private PreviewModule _previewSelection;
        public PreviewModule PreviewSelection
        {
            get => _previewSelection;

            set
            {
                _previewSelection = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MovePreview));
            }
        }


        public List<KeyValuePair<PreviewModule, string>> PreviewSelections
        {
            get
            {
                var previewSelections = new List<KeyValuePair<PreviewModule, string>>()
                {
                    new KeyValuePair<PreviewModule, string>(PreviewModule.Destination, RubberduckUI.MoveMember_Destination),
                    new KeyValuePair<PreviewModule, string>(PreviewModule.Source, $"{Model.Source.ModuleName}"),
                };
                return previewSelections;
            }
        }

        public string MoveableMembersLabel => string.Format(RubberduckUI.MoveMember_MoveMember_MemberListLabelFormat, Model.Source.ModuleName);

        public string DestinationModuleName
        {
            get => Model.Destination?.ModuleName ?? string.Empty;

            set
            {
                Model.ChangeDestination(value);
                OnPropertyChanged(nameof(IsExecutableMove));
                OnPropertyChanged(nameof(MovePreview));
            }
        }

        public Declaration DestinationModule
        {
            set
            {
                if (value is ModuleDeclaration modDeclaration)
                {
                    Model.ChangeDestination(modDeclaration);
                    OnPropertyChanged(nameof(IsExecutableMove));
                    OnPropertyChanged(nameof(MovePreview));
                }
            }
            get
            {
                if (Model.Destination.IsExistingModule(out var module))
                {
                    return module;
                }
                return null;
            }
        }

        public IEnumerable<KeyValuePair<Declaration, string>> DestinationModules
            => Modules(DeclarationType.ProceduralModule).Where(mod => mod.Key != Model.Source.Module);

        private IEnumerable<KeyValuePair<Declaration, string>> Modules(DeclarationType typeFlag)
        {
            return _declarationFinderProvider.DeclarationFinder.AllUserDeclarations
                            .Where(ud => ud.DeclarationType.HasFlag(typeFlag))
                            .OrderBy(ud => ud.IdentifierName)
                            .Select(mod => new KeyValuePair<Declaration, string>(mod, mod.IdentifierName));
        }

        public string DestinationSelectionLabel => string.Format(RubberduckUI.MoveMember_MoveMember_DestinationSelectionLabelFormat, LocalizedTypeDisplay(ComponentType.StandardModule));

        public string SourceModuleLabel => string.Format(RubberduckUI.MoveMember_MoveMember_SourceModuleLabelFormat, LocalizedTypeDisplay(Model.Source.ComponentType));

        private string LocalizedTypeDisplay(ComponentType componentType)
        {
            switch (componentType)
            {
                case ComponentType.ClassModule:
                    return RubberduckUI.ResourceManager.GetString("ComponentType_ClassModule", CultureInfo.CurrentUICulture);
                case ComponentType.UserForm:
                    return RubberduckUI.ResourceManager.GetString("ComponentType_UserForm", CultureInfo.CurrentUICulture);
                case ComponentType.StandardModule:
                    return RubberduckUI.ResourceManager.GetString("ComponentType_StandardModule", CultureInfo.CurrentUICulture);
                default:
                    return string.Empty;
            }
        }

        private ObservableCollection<MoveableMemberSetViewModel> _moveableMembers;
        public ObservableCollection<MoveableMemberSetViewModel> MoveCandidates
        {
            get
            {
                Debug.Assert(Model.Source.Module != null);
                _moveableMembers = new ObservableCollection<MoveableMemberSetViewModel>(OrderedMemberSets());
                return _moveableMembers;
            }
        }

        private IEnumerable<MoveableMemberSetViewModel> OrderedMemberSets()
        {
            return _moveableMemberSetViewModels
                .OrderByDescending(mm => mm.IsSelected)
                .ThenByDescending(mm => mm.IsPublicMember)
                .ThenByDescending(mm => mm.IsPrivateMember)
                .ThenByDescending(mm => mm.IsPublicConstant)
                .ThenByDescending(mm => mm.IsPublicField)
                .ThenByDescending(mm => mm.IsPrivateConstant)
                .ThenByDescending(mm => mm.ToDisplayString);
        }

        public void RefreshPreview(MoveableMemberSetViewModel selected)
        {
            OnPropertyChanged(nameof(MoveCandidates));
            OnPropertyChanged(nameof(IsExecutableMove));
            OnPropertyChanged(nameof(MovePreview));
        }

        public string MovePreview
        {
            get
            {
                var endpointToPreview = _previewSelection.Equals(PreviewModule.Destination)
                    ? Model.Destination as IMoveMemberEndpoint
                    : Model.Source as IMoveMemberEndpoint;

                return _previewerFactory.Create(endpointToPreview).PreviewMove(Model);
             }
        }

        private void SetAllSelections(bool value)
        {
            foreach (var item in MoveCandidates)
            {
                item.IsSelected = value;
            }
        }

        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }

        private bool AreVBAEquivalent(string identifier1, string identifier2)
        {
            return identifier1.Equals(identifier2, StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
