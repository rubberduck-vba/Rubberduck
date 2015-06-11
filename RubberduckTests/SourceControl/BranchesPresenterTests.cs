using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class BranchesPresenterTests
    {
        private Mock<ISourceControlProvider> _provider;
        private Mock<IBranchesView> _view;
        private Branch _intialBranch;
        private List<IBranch> _branches;
        private BranchesPresenter _presenter;
        private Mock<ICreateBranchView> _createView;
        private Mock<IMergeView> _mergeView;

        [TestInitialize]
        public void IntializeFixtures()
        {
            _provider = new Mock<ISourceControlProvider>();
            _view = new Mock<IBranchesView>();
            _createView = new Mock<ICreateBranchView>();
            _mergeView = new Mock<IMergeView>();

            _intialBranch = new Branch("master", "refs/Heads/master", false, true);

            //todo: create more realistic list of branches. Include `HEAD` so that we can ensure it gets excluded.

            _branches = new List<IBranch>()
            {
                _intialBranch,
                new Branch("dev", "ref/Heads/dev",isRemote: false, isCurrentHead:false),
                new Branch("origin/master", "refs/remotes/origin/master", true, true),
                new Branch("origin/HEAD", "refs/remotes/origin/HEAD", true, false)
            };

            _provider.SetupGet(git => git.Branches).Returns(_branches);
            _provider.SetupGet(git => git.CurrentBranch).Returns(_intialBranch);

            _presenter = new BranchesPresenter(_view.Object, _createView.Object, _mergeView.Object, _provider.Object);
        }

        [TestMethod]
        public void SelectedBranchShouldBeCurrentBranchAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.Current);            

            //act
            _presenter.RefreshView();

            //assert
            Assert.AreEqual(_provider.Object.CurrentBranch.Name, _view.Object.Current);
        }

        [TestMethod]
        public void PublishedBranchesAreListedAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.Published);

            //act
            _presenter.RefreshView();

            //assert
            var expected = new List<string>() {"master"};
            CollectionAssert.AreEqual(expected, _view.Object.Published.ToList());
        }

        [TestMethod]
        public void UnPublishedBranchesAreListedAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.Unpublished);

            //act
            _presenter.RefreshView();

            //assert
            var expected = new List<string>() {"dev"};
            CollectionAssert.AreEqual(expected, _view.Object.Unpublished.ToList());
        }

        [TestMethod]
        public void OnlyLocalBranchesInBranches()
        {
            //arrange 
            _view.SetupProperty(v => v.Local);

            //act
            _presenter.RefreshView();

            //assert
            var expected = new List<string>() {"master", "dev"};
            CollectionAssert.AreEquivalent(expected, _view.Object.Local.ToList());
        }

        [TestMethod]
        public void HeadIsNotIncludedInPublishedBranches()
        {
            //arrange
            _view.SetupProperty(v => v.Published);

            //act 
            _presenter.RefreshView();

            //assert
            CollectionAssert.DoesNotContain(_view.Object.Published.ToList(), "HEAD");
        }

        [TestMethod]
        public void CreateBranchViewIsShownOnCreateBranch()
        {
            //arrange
            _view.SetupProperty(v => v.Local, new List<string>());

            //act
            _view.Raise(v => v.CreateBranch += null, new EventArgs());

            //Assert
            _createView.Verify(c => c.Show(), Times.Once());
        }

        [TestMethod]
        public void CreateBranchViewBlocksBranchWithExistingName()
        {
            //arrange
            var branchName = "master";
            var branches = new List<string>() { branchName };

            _view.SetupProperty(v => v.Local, branches);
            _createView.SetupProperty(c => c.UserInputText, branchName);
            _createView.SetupProperty(c => c.OkButtonEnabled);

            //act
            _createView.Raise(c => c.UserInputTextChanged += null, new EventArgs());

            //Assert
            Assert.IsFalse(_createView.Object.OkButtonEnabled);
        }

        [TestMethod]
        public void CreateBranchViewAllowsValidBranchName()
        {
            //arrange
            var existingBranchName = "master";
            var newBranchName = "bugBranch";
            var branches = new List<string>() { existingBranchName };

            _view.SetupProperty(v => v.Local, branches);
            _createView.SetupProperty(c => c.UserInputText, newBranchName);
            _createView.SetupProperty(c => c.OkButtonEnabled);

            //act
            _createView.Raise(c => c.UserInputTextChanged += null, new EventArgs());

            //Assert
            Assert.IsTrue(_createView.Object.OkButtonEnabled);
        }

        [TestMethod]
        public void CreateBranchViewBlocksNameWithWhitespace()
        {
            //arrange
            var branchName = "my master";
            var branches = new List<string>() { branchName };

            _view.SetupProperty(v => v.Local, branches);
            _createView.SetupProperty(c => c.UserInputText, branchName);
            _createView.SetupProperty(c => c.OkButtonEnabled);

            //act
            _createView.Raise(c => c.UserInputTextChanged += null, new EventArgs());

            //Assert
            Assert.IsFalse(_createView.Object.OkButtonEnabled);
        }

        [TestMethod]
        public void CreateBranchViewIsNotShownWhenLocal_IsNull()
        {
            //arrange
            //act
            _view.Raise(v => v.CreateBranch += null, new EventArgs());

            //Assert
            _createView.Verify(c => c.Show(), Times.Never());
        }

        [TestMethod]
        public void ProviderCallsCreateBranchOnCreateBranchConfirm()
        {
            //arrange
            var expected = "testBranch";

            //act
            _createView.Raise(c => c.Confirm += null ,new BranchCreateArgs(expected));

            //assert
            _provider.Verify(git => git.CreateBranch(It.Is<string>(s => s == expected)));
        }

        [TestMethod]
        public void CreateBranchViewIshiddenAfterSubmit()
        {
            //arrange
            _createView.SetupProperty(c => c.UserInputText, "test");

            //act
            _createView.Raise(c => c.Confirm += null, new BranchCreateArgs(_createView.Object.UserInputText));

            //assert
            _createView.Verify(c => c.Hide(), Times.Once);
        }

        [TestMethod]
        public void CreateBranchUserInputIsClearedAfterSubmit()
        {
            //arrange
            _createView.SetupProperty(c => c.UserInputText, "test");

            //act
            _createView.Raise(c => c.Confirm += null, new BranchCreateArgs("test"));

            //assert
            Assert.AreEqual(string.Empty, _createView.Object.UserInputText);
        }

        [TestMethod]
        public void MergeViewIsShownOnMergeClick()
        {
            //arrange
            _view.SetupProperty(v => v.Local, new List<string>());

            //act
            _view.Raise(v => v.Merge += null, new EventArgs());

            //assert
            _mergeView.Verify(m => m.Show(), Times.Once);
        }

        [TestMethod]
        public void MergeViewIsNotShownWhenLocal_IsNull()
        {
            //arrange
            _view.SetupProperty(v => v.Local); //no default value, so v.Local is null

            //act
            _view.Raise(v => v.Merge += null, new EventArgs());

            //assert
            _mergeView.Verify(m => m.Show(), Times.Never);
        }

        [TestMethod]
        public void MergeViewSourceBranchesAreSetBeforeShowing()
        {
            //arrange
            _mergeView.SetupProperty(m => m.SourceSelectorData);
            _view.SetupProperty(v => v.Local, new List<string>() {"master", "dev"});

            //act
            _view.Raise(v => v.Merge += null, new EventArgs());

            //assert
            CollectionAssert.AreEqual(_view.Object.Local.ToList(), _mergeView.Object.SourceSelectorData.ToList());
        }

        [TestMethod]
        public void MergeViewSelectedSourceBranchIsCurrentBranch()
        {
            //arrange
            _mergeView.SetupProperty(m => m.SourceSelectorData);
            _view.SetupProperty(v => v.Local, new List<string>() { "master", "dev" });

            _mergeView.SetupProperty(m => m.SelectedSourceBranch);

            //act
            _view.Raise(v => v.Merge += null, new EventArgs());

            //assert 
            Assert.AreEqual(_intialBranch.Name, _mergeView.Object.SelectedSourceBranch);
        }

        [TestMethod]
        public void MergeViewDestinationBranchesAreSetBeforeShowing()
        {
            //arrange
            _mergeView.SetupProperty(m => m.DestinationSelectorData);
            _view.SetupProperty(v => v.Local, new List<string>() { "master", "dev" });

            //act
            _view.Raise(v => v.Merge += null, new EventArgs());

            //assert
            CollectionAssert.AreEqual(_view.Object.Local.ToList(), _mergeView.Object.DestinationSelectorData.ToList());
        }

        [TestMethod]
        public void ProviderMergesOnMergeViewSubmit()
        {
            //arrange
            _mergeView.SetupProperty(m => m.SelectedSourceBranch, "dev");
            _mergeView.SetupProperty(m => m.SelectedDestinationBranch, "master");

            //act
            _mergeView.Raise(m => m.Confirm += null, new EventArgs());

            //assert
            _provider.Verify(git => git.Merge("dev", "master"));
        }

        [TestMethod]
        public void MergeViewIsHiddenOnSuccessfulMerge()
        {
            //arrange
            _mergeView.SetupProperty(m => m.SelectedSourceBranch, "dev");
            _mergeView.SetupProperty(m => m.SelectedDestinationBranch, "master");

            //act
            _mergeView.Raise(m => m.Confirm += null, new EventArgs());

            //assert
            _mergeView.Verify(m => m.Hide());
        }

        [TestMethod]
        public void MergeViewIsHiddenOnCancel()
        {
            //act
            _mergeView.Raise(m => m.Cancel += null, new EventArgs());

            //assert
            _mergeView.Verify(m => m.Hide());
        }

        [TestMethod]
        public void MergeStatusHiddenWhenViewIsFirstShown()
        {
            //arrange
            _mergeView.SetupProperty(m => m.StatusTextVisible, false);

            //act
            _mergeView.Object.Show();

            //assert
            Assert.IsFalse(_mergeView.Object.StatusTextVisible, "Merge Status Is Visible");
        }

        [TestMethod]
        public void MergeStatusIsUnknownWhenViewIsFirstShown()
        {
            //arrange
            _mergeView.SetupProperty(m => m.Status);
            
            //act
            _mergeView.Object.Show();

            //assert
            Assert.AreEqual(MergeStatus.Unknown, _mergeView.Object.Status);
        }

        [TestMethod]
        public void MergeStatusIsVisibleOnSuccess()
        {
            //arrange
            _mergeView.SetupProperty(m => m.StatusTextVisible, false);
            _mergeView.SetupProperty(m => m.Status);

            //act
            _mergeView.Object.Status = MergeStatus.Success;
            _mergeView.Raise(m => m.MergeStatusChanged += null, EventArgs.Empty);
            
            //Assert
            Assert.IsTrue(_mergeView.Object.StatusTextVisible, "Merge Status Is Not Visible");
        }

        [TestMethod]
        public void MergeStatusIsVisibleOnFailure()
        {
            //arrange
            _mergeView.SetupProperty(m => m.StatusTextVisible, false);
            _mergeView.SetupProperty(m => m.Status);

            //act
            _mergeView.Object.Status = MergeStatus.Failure;
            _mergeView.Raise(m => m.MergeStatusChanged += null, EventArgs.Empty);

            //Assert
            Assert.IsTrue(_mergeView.Object.StatusTextVisible, "Merge Status Is Not Visible"); 
        }

        [TestMethod]
        public void MergeStatusTextIsEmptiedWhenStatusIsChangedToUnknown()
        {
            //arrange
            _mergeView.SetupProperty(m => m.StatusText, "Some Text");
            _mergeView.SetupProperty(m => m.Status, MergeStatus.Failure);

            //act
            _mergeView.Object.Status = MergeStatus.Unknown;
            _mergeView.Raise(m => m.MergeStatusChanged +=null, EventArgs.Empty);

            //assert
            Assert.AreEqual(String.Empty, _mergeView.Object.StatusText);
        }

        [TestMethod]
        [ExpectedException(typeof(SourceControlException))]
        public void MergeStatusSetToFailIfProviderThrowsException()
        {
            //arrange
            _mergeView.SetupProperty(m => m.Status, MergeStatus.Unknown);
            _provider.Setup(git => git.Merge(It.IsAny<string>(), It.IsAny<string>()))
                .Throws(new SourceControlException());

            //act
            _provider.Object.Merge("dev", "master");

            //assert
            Assert.AreEqual(MergeStatus.Failure, _mergeView.Object.Status);
        }

        [TestMethod]
        public void ChangingSelectedBranchChecksOutThatBranch()
        {
            //arrange
            _view.SetupProperty(v => v.Current, "master");
            _provider.Setup(git => git.Checkout(It.IsAny<string>()));

            //act
            _view.Object.Current = "dev";
            _view.Raise(v => v.SelectedBranchChanged+=null, EventArgs.Empty);

            //assert
            _provider.Verify(git => git.Checkout("dev"));
        }

        [TestMethod]
        public void RefreshingViewShouldNotCheckoutBranch()
        {
            //arrange
            _view.SetupProperty(v => v.Current, "master");
            _provider.Setup(git => git.Checkout(It.IsAny<string>()));

            //act
            _presenter.RefreshView();

            //assert
            _provider.Verify(git => git.Checkout(It.IsAny<string>()),Times.Never);
        }
    }
}
