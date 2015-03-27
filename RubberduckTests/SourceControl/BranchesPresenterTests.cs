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

            _branches = new List<IBranch>()
            {
                _intialBranch,
                new Branch("dev", "ref/Heads/dev", false, false),
                new Branch("origin/master", "refs/remotes/origin/master", true, true),
            };

            _provider.SetupGet(git => git.Branches).Returns(_branches);
            _provider.SetupGet(git => git.CurrentBranch).Returns(_intialBranch);
        }

        [TestMethod]
        public void SelectedBranchShouldBeCurrentBranchAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.CurrentBranch);            
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            Assert.AreEqual(_provider.Object.CurrentBranch.Name, _view.Object.CurrentBranch);
        }

        [TestMethod]
        public void PublishedBranchesAreListedAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.PublishedBranches);
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            var expected = new List<string>() {"master"};
            CollectionAssert.AreEqual(expected, _view.Object.PublishedBranches.ToList());
        }

        [TestMethod]
        public void UnPublishedBranchesAreListedAfterRefresh()
        {
            //arrange
            _view.SetupProperty(v => v.UnpublishedBranches);
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //act
            presenter.RefreshView();

            //assert
            var expected = new List<string>() {"dev"};
            CollectionAssert.AreEqual(expected, _view.Object.UnpublishedBranches.ToList());
        }


    }
}
