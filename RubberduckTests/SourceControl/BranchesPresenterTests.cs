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

            var branches = new List<string>() { "master", "dev" };
            _provider.SetupGet(git => git.Branches).Returns(branches);
            _provider.SetupGet(git => git.CurrentBranch).Returns("dev");

            _view.SetupProperty(v => v.CurrentBranch);

            //act
            var presenter = new BranchesPresenter(_provider.Object, _view.Object);

            //assert
            Assert.AreEqual(_provider.Object.CurrentBranch, _view.Object.CurrentBranch);
        }

    }
}
