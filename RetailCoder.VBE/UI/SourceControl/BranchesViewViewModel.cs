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
    public class BranchesViewViewModel : ViewModelBase, IControlViewModel
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public BranchesViewViewModel()
        {
            _newBranchCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CreateBranch(), _ => Provider != null);
            _mergeBranchCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => MergeBranch(), _ => Provider != null);

            _createBranchOkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CreateBranchOk(), _ => !IsNotValidBranchName);
            _createBranchCancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CreateBranchCancel());

            _mergeBranchesOkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => MergeBranchOk(), _ => SourceBranch != DestinationBranch);
            _mergeBranchesCancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => MergeBranchCancel());

            _deleteBranchToolbarButtonCommand =
                new DelegateCommand(LogManager.GetCurrentClassLogger(), isBranchPublished => DeleteBranch(bool.Parse((string) isBranchPublished)),
                    isBranchPublished => CanDeleteBranch(bool.Parse((string)isBranchPublished)));

            _publishBranchToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => PublishBranch(), _ => !string.IsNullOrEmpty(CurrentUnpublishedBranch));
            _unpublishBranchToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UnpublishBranch(), _ => !string.IsNullOrEmpty(CurrentPublishedBranch));
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
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

        public SourceControlTab Tab { get { return SourceControlTab.Branches; } }

        public IEnumerable<string> Branches
        {
            get
            {
                return Provider == null
                  ? Enumerable.Empty<string>()
                  : Provider.Branches.Select(b => b.Name);
            }
        }

        public IEnumerable<string> LocalBranches
        {
            get
            {
                return Provider == null
                    ? Enumerable.Empty<string>()
                    : Provider.Branches.Where(b => !b.IsRemote).Select(b => b.Name);
            }
        }

        public IEnumerable<string> PublishedBranches
        {
            get
            {
                return Provider == null
                    ? Enumerable.Empty<string>()
                    : Provider.Branches.Where(b => !b.IsRemote && !string.IsNullOrEmpty(b.TrackingName)).Select(b => b.Name);
            }
        }

        public IEnumerable<string> UnpublishedBranches
        {
            get
            {
                return Provider == null
                    ? Enumerable.Empty<string>()
                    : Provider.Branches.Where(b => !b.IsRemote && string.IsNullOrEmpty(b.TrackingName)).Select(b => b.Name);
            }
        }

        private string _currentBranch;
        public string CurrentBranch
        {
            get { return _currentBranch; }
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
            get { return _currentPublishedBranch; }
            set
            {
                _currentPublishedBranch = value;
                OnPropertyChanged();
            }
        }

        private string _currentUnpublishedBranch;
        public string CurrentUnpublishedBranch
        {
            get { return _currentUnpublishedBranch; }
            set
            {
                _currentUnpublishedBranch = value;
                OnPropertyChanged();
            }
        }

        private bool _displayCreateBranchGrid;
        public bool DisplayCreateBranchGrid
        {
            get { return _displayCreateBranchGrid; }
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
            get { return _createBranchSource; }
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
            get { return _newBranchName; }
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

        public bool IsNotValidBranchName
        {
            get { return string.IsNullOrEmpty(NewBranchName) || !ValidBranchNameRegex.IsMatch(NewBranchName); }
        }

        private bool _displayMergeBranchesGrid;
        public bool DisplayMergeBranchesGrid
        {
            get { return _displayMergeBranchesGrid; }
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
            get { return _sourceBranch; }
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
            get { return _destinationBranch; }
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
            Logger.Trace("Creating branch {0}", NewBranchName);
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
            Logger.Trace("Merging branch {0} into branch {1}", SourceBranch, DestinationBranch);

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
            Logger.Trace("Deleting {0}published branch {1}", isBranchPublished ? "" : "un",
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
            Logger.Trace("Publishing branch {0}", CurrentUnpublishedBranch);
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
            Logger.Trace("Unpublishing branch {0}", CurrentPublishedBranch);
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

        private readonly CommandBase _newBranchCommand;
        public CommandBase NewBranchCommand
        {
            get
            {
                return _newBranchCommand;
            }
        }

        private readonly CommandBase _mergeBranchCommand;
        public CommandBase MergeBranchCommand
        {
            get
            {
                return _mergeBranchCommand;
            }
        }

        private readonly CommandBase _createBranchOkButtonCommand;
        public CommandBase CreateBranchOkButtonCommand
        {
            get
            {
                return _createBranchOkButtonCommand;
            }
        }

        private readonly CommandBase _createBranchCancelButtonCommand;
        public CommandBase CreateBranchCancelButtonCommand
        {
            get
            {
                return _createBranchCancelButtonCommand;
            }
        }

        private readonly CommandBase _mergeBranchesOkButtonCommand;
        public CommandBase MergeBranchesOkButtonCommand
        {
            get
            {
                return _mergeBranchesOkButtonCommand;
            }
        }

        private readonly CommandBase _mergeBranchesCancelButtonCommand;
        public CommandBase MergeBranchesCancelButtonCommand
        {
            get
            {
                return _mergeBranchesCancelButtonCommand;
            }
        }

        private readonly CommandBase _deleteBranchToolbarButtonCommand;
        public CommandBase DeleteBranchToolbarButtonCommand
        {
            get
            {
                return _deleteBranchToolbarButtonCommand;
            }
        }

        private readonly CommandBase _publishBranchToolbarButtonCommand;
        public CommandBase PublishBranchToolbarButtonCommand
        {
            get { return _publishBranchToolbarButtonCommand; }
        }

        private readonly CommandBase _unpublishBranchToolbarButtonCommand;
        public CommandBase UnpublishBranchToolbarButtonCommand
        {
            get { return _unpublishBranchToolbarButtonCommand; }
        }

        public event EventHandler<ErrorEventArgs> ErrorThrown;
        private void RaiseErrorEvent(string message, Exception innerException, NotificationType notificationType)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(message, innerException, notificationType));
            }
        }

        private void RaiseErrorEvent(string title, string message, NotificationType notificationType)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(title, message, notificationType));
            }
        }
    }
}
