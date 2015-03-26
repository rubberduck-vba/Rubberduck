using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;
using Moq;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class BranchesPresenterTests
    {
        [TestMethod]
        public void SelectedBranchShouldBeCurrentBranchAfterRefresh()
        {

            //arrange
            var _provider = new Mock<ISourceControlProvider>();
            var _view = new Mock<IBranchesView>();

            var expectedBranch = new Branch("dev", "ref/Heads/dev", false, false);

            var branches = new List<IBranch>()
            {
                new Branch("master", "refs/Heads/master", false, true),
                expectedBranch,
                new Branch("origin/master", "refs/remotes/origin/master", true, true),
                new Branch("origin/dev", "refs/remotes/origin/dev", true, false)
            };

            _provider.SetupGet(git => git.Branches).Returns(branches);
            _provider.SetupGet(git => git.CurrentBranch).Returns(expectedBranch);

            _view.SetupProperty(v => v.CurrentBranch);

            //act
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //assert
            Assert.AreEqual(_provider.Object.CurrentBranch.Name, _view.Object.CurrentBranch);
        }

    }
}
