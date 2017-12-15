using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NLog;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
// ReSharper disable ExplicitCallerInfoArgument

namespace Rubberduck.UI.SourceControl
{
    public class BranchesPanelViewModel : ViewModelBase, IControlViewModel
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public BranchesPanelViewModel()
        {
            NewBranchCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CreateBranch(), _ => Provider != null);
            MergeBranchCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => MergeBranch(), _ => Provider != null);

            CreateBranchOkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CreateBranchOk(), _ => !IsNotValidBranchName);
            CreateBranchCancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CreateBranchCancel());

            MergeBranchesOkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => MergeBranchOk(), _ => SourceBranch != DestinationBranch);
            MergeBranchesCancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => MergeBranchCancel());

            DeleteBranchToolbarButtonCommand =
                new DelegateCommand(LogManager.GetCurrentClassLogger(), isBranchPublished => DeleteBranch(bool.Parse((string) isBranchPublished)),
                    isBranchPublished => CanDeleteBranch(bool.Parse((string)isBranchPublished)));

            PublishBranchToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => PublishBranch(), _ => !string.IsNullOrEmpty(CurrentUnpublishedBranch));
            UnpublishBranchToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UnpublishBranch(), _ => !string.IsNullOrEmpty(CurrentPublishedBranch));
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get => _provider;
            set
            {
                Logger.Trace("Provider changed");

                _provider = value;
                RefreshView();
            }
        }

        public void RefreshView()
        {
            Logger.Trace("Refreshing view");

            OnPropertyChanged("LocalBranches");
            OnPropertyChanged("PublishedBranches");
            OnPropertyChanged("UnpublishedBranches");
            OnPropertyChanged("Branches");

            CurrentBranch = Provider.CurrentBranch.Name;

            SourceBranch = null;
            DestinationBranch = CurrentBranch;
        }

        public void ResetView()
        {
            Logger.Trace("Resetting view");

            _provider = null;
            _currentBranch = string.Empty;
            SourceBranch = string.Empty;
            DestinationBranch = CurrentBranch;

            OnPropertyChanged("LocalBranches");
            OnPropertyChanged("PublishedBranches");
            OnPropertyChanged("UnpublishedBranches");
            OnPropertyChanged("Branches");
            OnPropertyChanged("CurrentBranch");
        }

        public SourceControlTab Tab => SourceControlTab.Branches;

        public IEnumerable<string> Branches
        {
            get
            {
                return Provider?.Branches.Select(b => b.Name) ?? Enumerable.Empty<string>();
            }
        }

        public IEnumerable<string> LocalBranches
        {
            get
            {
                return Provider?.Branches.Where(b => !b.IsRemote).Select(b => b.Name) 
                    ?? Enumerable.Empty<string>();
            }
        }

        public IEnumerable<string> PublishedBranches
        {
            get
            {
                return Provider?.Branches.Where(b => !b.IsRemote && !string.IsNullOrEmpty(b.TrackingName)).Select(b => b.Name) 
                    ?? Enumerable.Empty<string>();
            }
        }

        public IEnumerable<string> UnpublishedBranches
        {
            get
            {
                return Provider?.Branches.Where(b => !b.IsRemote && string.IsNullOrEmpty(b.TrackingName)).Select(b => b.Name)
                    ?? Enumerable.Empty<string>();
            }
        }

        private string _currentBranch;
        public string CurrentBranch
        {
            get => _currentBranch;
            set
            {
                if (_currentBranch != value)
                {
                    _currentBranch = value;
                    OnPropertyChanged();

                    CreateBranchSource = value;

                    if (Provider == null) { return; }

                    try
                    {
                        Provider.NotifyExternalFileChanges = false;
                        Provider.HandleVbeSinkEvents = false;
                        Provider.Checkout(_currentBranch);
                    }
                    catch (SourceControlException ex)
                    {
                        RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
                    }
                    catch
                    {
                        RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                            RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                        throw;
                    }
                    Provider.NotifyExternalFileChanges = true;
                    Provider.HandleVbeSinkEvents = true;
                }
            }
        }

        private string _currentPublishedBranch;
        public string CurrentPublishedBranch
        {
            get => _currentPublishedBranch;
            set
            {
                _currentPublishedBranch = value;
                OnPropertyChanged();
            }
        }

        private string _currentUnpublishedBranch;
        public string CurrentUnpublishedBranch
        {
            get => _currentUnpublishedBranch;
            set
            {
                _currentUnpublishedBranch = value;
                OnPropertyChanged();
            }
        }

        private bool _displayCreateBranchGrid;
        public bool DisplayCreateBranchGrid
        {
            get => _displayCreateBranchGrid;
            set
            {
                if (_displayCreateBranchGrid != value)
                {
                    _displayCreateBranchGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _createBranchSource;
        public string CreateBranchSource
        {
            get => _createBranchSource;
            set
            {
                if (_createBranchSource != value)
                {
                    _createBranchSource = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _newBranchName;
        public string NewBranchName
        {
            get => _newBranchName;
            set
            {
                if (_newBranchName != value)
                {
                    _newBranchName = value;
                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidBranchName");
                }
            }
        }

        // Courtesy of http://stackoverflow.com/a/12093994/4088852 - Assumes --allow-onelevel is set TODO: Verify provider honor that. 
        private static readonly Regex ValidBranchNameRegex = new Regex(@"^(?!@$|build-|/|.*([/.]\.|//|@\{|\\))[^\u0000-\u0037\u0177 ~^:?*[]+/?[^\u0000-\u0037\u0177 ~^:?*[]+(?<!\.lock|[/.])$");

        public bool IsNotValidBranchName => string.IsNullOrEmpty(NewBranchName) || !ValidBranchNameRegex.IsMatch(NewBranchName);

        private bool _displayMergeBranchesGrid;
        public bool DisplayMergeBranchesGrid
        {
            get => _displayMergeBranchesGrid;
            set
            {
                if (_displayMergeBranchesGrid != value)
                {
                    _displayMergeBranchesGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _sourceBranch;
        public string SourceBranch
        {
            get => _sourceBranch;
            set
            {
                if (_sourceBranch != value)
                {
                    _sourceBranch = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _destinationBranch;
        public string DestinationBranch
        {
            get => _destinationBranch;
            set
            {
                if (_destinationBranch != value)
                {
                    _destinationBranch = value;
                    OnPropertyChanged();
                }
            }
        }

        private void CreateBranch()
        {
            DisplayMergeBranchesGrid = false;

            DisplayCreateBranchGrid = true;
            NewBranchName = string.Empty;
        }

        private void MergeBranch()
        {
            DisplayCreateBranchGrid = false;

            DisplayMergeBranchesGrid = true;
        }

        private void CreateBranchOk()
        {
            Logger.Trace($"Creating branch {NewBranchName}");
            try
            {
                Provider.CreateBranch(CreateBranchSource, NewBranchName);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }

            DisplayCreateBranchGrid = false;
            NewBranchName = string.Empty;

            RefreshView();
        }

        private void CreateBranchCancel()
        {
            DisplayCreateBranchGrid = false;
            NewBranchName = string.Empty;
        }

        private void MergeBranchOk()
        {
            Logger.Trace($"Merging branch {SourceBranch} into branch {DestinationBranch}");

            Provider.NotifyExternalFileChanges = false;
            Provider.HandleVbeSinkEvents = false;

            try
            {
                Provider.Merge(SourceBranch, DestinationBranch);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
                Provider.NotifyExternalFileChanges = true;
                Provider.HandleVbeSinkEvents = true;
                return;
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                Provider.NotifyExternalFileChanges = true;
                Provider.HandleVbeSinkEvents = true;
                throw;
            }

            DisplayMergeBranchesGrid = false;
            RaiseErrorEvent(RubberduckUI.SourceControl_MergeStatus, string.Format(RubberduckUI.SourceControl_SuccessfulMerge, SourceBranch, DestinationBranch), NotificationType.Info);

            Provider.NotifyExternalFileChanges = true;
            Provider.HandleVbeSinkEvents = true;
        }

        private void MergeBranchCancel()
        {
            DisplayMergeBranchesGrid = false;
        }

        private void DeleteBranch(bool isBranchPublished)
        {
            Logger.Trace("Deleting {0}published branch {1}",
                isBranchPublished ? "" : "un",
                isBranchPublished ? CurrentPublishedBranch : CurrentUnpublishedBranch);
            try
            {
                Provider.DeleteBranch(isBranchPublished ? CurrentPublishedBranch : CurrentUnpublishedBranch);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }

            RefreshView();
        }
        
        private bool CanDeleteBranch(bool isBranchPublished)
        {
            return isBranchPublished
                ? !string.IsNullOrEmpty(CurrentPublishedBranch) && CurrentPublishedBranch != CurrentBranch
                : !string.IsNullOrEmpty(CurrentUnpublishedBranch) && CurrentUnpublishedBranch != CurrentBranch;
        }

        private void PublishBranch()
        {
            Logger.Trace($"Publishing branch {CurrentUnpublishedBranch}");
            try
            {
                Provider.Publish(CurrentUnpublishedBranch);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }

            RefreshView();
        }

        private void UnpublishBranch()
        {
            Logger.Trace($"Unpublishing branch {CurrentPublishedBranch}");
            try
            {
                Provider.Unpublish(Provider.Branches.First(b => b.Name == CurrentPublishedBranch).TrackingName);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }

            RefreshView();
        }

        public CommandBase NewBranchCommand { get; }

        public CommandBase MergeBranchCommand { get; }

        public CommandBase CreateBranchOkButtonCommand { get; }

        public CommandBase CreateBranchCancelButtonCommand { get; }

        public CommandBase MergeBranchesOkButtonCommand { get; }

        public CommandBase MergeBranchesCancelButtonCommand { get; }

        public CommandBase DeleteBranchToolbarButtonCommand { get; }

        public CommandBase PublishBranchToolbarButtonCommand { get; }

        public CommandBase UnpublishBranchToolbarButtonCommand { get; }

        public event EventHandler<ErrorEventArgs> ErrorThrown;
        private void RaiseErrorEvent(string message, Exception innerException, NotificationType notificationType)
        {
            ErrorThrown?.Invoke(this, new ErrorEventArgs(message, innerException, notificationType));
        }

        private void RaiseErrorEvent(string title, string message, NotificationType notificationType)
        {
            ErrorThrown?.Invoke(this, new ErrorEventArgs(title, message, notificationType));
        }
    }
}
