using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.MoveMember.Extensions;
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
        private List<MoveableMemberSetViewModel> _moveableMemberViewModels;
        private List<string> _existingModuleNames;
        public MoveMemberViewModel(MoveMemberModel model)
            : base(model)
        {
            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SetAllSelections(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SetAllSelections(false));

            var moveableMembers = model.MoveableMembers;

            _moveableMemberViewModels = new List<MoveableMemberSetViewModel>();
            foreach (var mm in moveableMembers)
            {
                if (mm.Member.IsVariable() && mm.Member.HasPrivateAccessibility())
                {
                    continue;
                }
                var moveableMemberViewModel = new MoveableMemberSetViewModel(this, mm);
                _moveableMemberViewModels.Add(moveableMemberViewModel);
            }
            _moveableMembers = new ObservableCollection<MoveableMemberSetViewModel>(OrderedMemberSets());

            _existingModuleNames = Model.DeclarationFinderProvider.DeclarationFinder.AllUserDeclarations
                .Where(ud => ud.DeclarationType.HasFlag(DeclarationType.Module))
                .Select(m => m.IdentifierName).ToList();

            _previewSelection = PreviewModule.Destination;
            _destinationNameFailureCriteria = string.Empty;
        }

        public string RefactorName => MoveMemberResources.RefactorName;

        public string Instructions => MoveMemberResources.Instructions;

        public bool IsExecutableMove
        {
            get
            {
                var result = false;
                if (IsValidModuleName && MoveMemberObjectsFactory.TryCreateStrategy(Model, out var strategy))
                {
                    if (strategy.IsExecutableModel(Model, out _nonExecutableMessage))
                    {
                        result = true;
                    }
                }
                OnPropertyChanged(nameof(IsValidModuleName));
                OnPropertyChanged(nameof(DestinationNameFailureCriteria));
                return result;
            }
        }

        private string _nonExecutableMessage;
        public string NonExecutableMessage => _nonExecutableMessage;

        private string _destinationNameFailureCriteria;
        public string DestinationNameFailureCriteria
        {
            get
            {
                if (!string.IsNullOrEmpty(_destinationNameFailureCriteria))
                {
                    return _destinationNameFailureCriteria;
                }
                if (!string.IsNullOrEmpty(_nonExecutableMessage))
                {
                    return _nonExecutableMessage;
                }
                return string.Empty;
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

            if (Model.Source.Module.ProjectName.IsEquivalentVBAIdentifierTo(DestinationModuleName))
            {
                failCriteria = MoveMemberResources.ModuleMatchesProjectNameFailMsg; 
                return false;
            }

            if (Model.Source.ModuleName.IsEquivalentVBAIdentifierTo(DestinationModuleName))
            {
                failCriteria = MoveMemberResources.SourceAndDestinationModuleNameMatcheFailMsg;
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
                    new KeyValuePair<PreviewModule, string>(PreviewModule.Destination, MoveMemberResources.MoveMember_Destination),
                    new KeyValuePair<PreviewModule, string>(PreviewModule.Source, $"{Model.Source.ModuleName}"),
                };
                return previewSelections;
            }
        }

        public string MoveableMembersLabel => string.Format(MoveMemberResources.MoveMember_MemberListLabelFormat, Model.Source.ModuleName);

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
                Model.ChangeDestination(value);
                OnPropertyChanged(nameof(IsExecutableMove));
                OnPropertyChanged(nameof(MovePreview));
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

        private IEnumerable<KeyValuePair<Declaration, string>> Modules(Enum typeFlag)
        {
            return Model.DeclarationFinderProvider.DeclarationFinder.AllUserDeclarations
                            .Where(ud => ud.DeclarationType.HasFlag(typeFlag))
                            .OrderBy(ud => ud.IdentifierName)
                            .Select(mod => new KeyValuePair<Declaration, string>(mod, mod.IdentifierName));
        }

        public string DestinationSelectionLabel => string.Format(MoveMemberResources.MoveMember_DestinationSelectionLabelFormat, LocalizedTypeDisplay(ComponentType.StandardModule));

        public string SourceModuleLabel => string.Format(MoveMemberResources.MoveMember_SourceModuleLabelFormat, LocalizedTypeDisplay(Model.Source.ComponentType));

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
            return _moveableMemberViewModels.OrderByDescending(mm => mm.IsSelected).ThenBy(mm => mm.ToDisplayString);
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
                return Model.PreviewModuleContent(_previewSelection);
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
    }
}
