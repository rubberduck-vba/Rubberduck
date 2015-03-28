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

        [TestInitialize]
        public void IntializeFixtures()
        {
            //arrange
            _provider = new Mock<ISourceControlProvider>();
            _view = new Mock<IBranchesView>();

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
        }

        [TestMethod]
        public void SelectedBranchShouldBeCurrentBranchAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.Current);            
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            Assert.AreEqual(_provider.Object.CurrentBranch.Name, _view.Object.Current);
        }

        [TestMethod]
        public void PublishedBranchesAreListedAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.Published);
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            var expected = new List<string>() {"master"};
            CollectionAssert.AreEqual(expected, _view.Object.Published.ToList());
        }

        [TestMethod]
        public void UnPublishedBranchesAreListedAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.Unpublished);
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            var expected = new List<string>() {"dev"};
            CollectionAssert.AreEqual(expected, _view.Object.Unpublished.ToList());
        }

        [TestMethod]
        public void OnlyLocalBranchesInBranches()
        {
            //arrange 
            _view.SetupProperty(v => v.Local);
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            var expected = new List<string>() {"master", "dev"};
            CollectionAssert.AreEquivalent(expected, _view.Object.Local.ToList());
        }

        [TestMethod]
        public void HeadIsNotIncludedInPublishedBranches()
        {
            //arrange
            _view.SetupProperty(v => v.Published);
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act 
            presenter.RefreshView();

            //assert
            CollectionAssert.DoesNotContain(_view.Object.Published.ToList(), "HEAD");
        }


    }
}
