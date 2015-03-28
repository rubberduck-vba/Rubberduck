using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;
using Moq;

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

        [TestInitialize]
        public void IntializeFixtures()
        {
            _provider = new Mock<ISourceControlProvider>();
            _view = new Mock<IBranchesView>();
            _createView = new Mock<ICreateBranchView>();

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

            _presenter = new BranchesPresenter(_provider.Object, _view.Object, _createView.Object);
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
            //act
            _view.Raise(v => v.CreateBranch += null, new EventArgs());

            //Assert
            _createView.Verify(c => c.Show(), Times.Once());
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
    }
}
