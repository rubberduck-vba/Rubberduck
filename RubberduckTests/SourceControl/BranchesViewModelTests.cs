using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class BranchesViewModelTests
    {
        private Mock<ISourceControlProvider> _provider;
        private Branch _intialBranch;
        private List<IBranch> _branches;

        [TestCategory("SourceControl")]
        [TestInitialize]
        public void IntializeFixtures()
        {
            _provider = new Mock<ISourceControlProvider>();

            var masterRemote = new Mock<LibGit2Sharp.Branch>();
            masterRemote.SetupGet(git => git.Tip).Returns(new Mock<LibGit2Sharp.Commit>().Object);
            masterRemote.SetupGet(git => git.FriendlyName).Returns("master");

            _intialBranch = new Branch("master", "refs/Heads/master", false, true, masterRemote.Object);

            //todo: create more realistic list of branches. Include `HEAD` so that we can ensure it gets excluded.

            _branches = new List<IBranch>
            {
                _intialBranch,
                new Branch("dev", "ref/Heads/dev", false, false, null),
                new Branch("origin/master", "refs/remotes/origin/master", true, true, null),
                new Branch("origin/HEAD", "refs/remotes/origin/HEAD", true, false, null)
            };

            _provider.SetupGet(git => git.Branches).Returns(_branches);
            _provider.SetupGet(git => git.CurrentBranch).Returns(_intialBranch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void SelectedBranchShouldBeCurrentBranchAfterRefresh()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.SetupGet(git => git.CurrentBranch).Returns(_branches[1]);

            vm.RefreshView();

            Assert.AreEqual(_provider.Object.CurrentBranch.Name, vm.CurrentBranch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void PublishedBranchesAreListed()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            var expected = new List<string> { "master" };
            CollectionAssert.AreEqual(expected, vm.PublishedBranches.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnPublishedBranchesAreListed()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            var expected = new List<string> { "dev" };
            CollectionAssert.AreEqual(expected, vm.UnpublishedBranches.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnlyLocalBranchesInBranches()
        { 
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            var expected = new List<string> { "master", "dev" };
            CollectionAssert.AreEquivalent(expected, vm.LocalBranches.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void HeadIsNotIncludedInPublishedBranches()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            CollectionAssert.DoesNotContain(vm.PublishedBranches.ToList(), "HEAD");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void DeleteBranchDisabled_BranchIsActive()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "master"
            };

            Assert.IsFalse(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.TrueString));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void DeleteBranchEnabled_BranchIsNotActive()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "bugbranch"
            };

            Assert.IsTrue(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.TrueString));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void DeleteBranch_BranchIsNotActive_BranchIsRemoved()
        {
            var firstBranchName = "master";
            var secondBranchName = "bugBranch";

            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            Assert.IsFalse(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.FalseString));

            _provider.SetupGet(p => p.Branches).Returns(
                new List<IBranch>
                {
                    new Branch(firstBranchName, "ref/Heads/" + firstBranchName, false, true, null),
                    new Branch(secondBranchName, "ref/Heads/" + secondBranchName, false, false, null)
                });

            vm.CurrentPublishedBranch = firstBranchName;
            vm.CurrentUnpublishedBranch = secondBranchName;

            _provider.Setup(p => p.DeleteBranch(It.IsAny<string>()));

            vm.DeleteBranchToolbarButtonCommand.Execute(bool.FalseString);

            _provider.Verify(p => p.DeleteBranch(secondBranchName));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIsShownOnCreateBranch()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            vm.NewBranchCommand.Execute(null);

            Assert.IsTrue(vm.DisplayCreateBranchGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeBranchViewIsShownOnCreateBranch()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            vm.NewBranchCommand.Execute(null);

            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_BranchExists()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute("master"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_ValidBranchName()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch"
            };

            Assert.IsTrue(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsSpace()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsTwoConsecutiveDots()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug..branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsTilde()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug~branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsCaret()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug^branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsColon()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug:branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsQuestionMark()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug?branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsAsteriks()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug*branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsOpenBracket()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug[branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsTwoConsecutiveSlashes()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug//branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameStartsWithSlash()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "/bugBranch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameEndsWithSlash()
        {            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch/"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameEndsWithDot()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch."
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameIsAtSign()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "@"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsConsecutiveAtSignAndOpenBrace()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug@{branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsBackslash()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug\\branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsSlashSectionStartingWithDot()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug/.branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsSlashSectionEndingWithDotlock()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug/branch.lock"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIsNotShownWhenLocal_IsNull()
        {
            var vm = new BranchesViewViewModel();

            Assert.IsFalse(vm.NewBranchCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderCallsCreateBranchOnCreateBranchConfirm()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch"
            };

            vm.CreateBranchOkButtonCommand.Execute(null);

            _provider.Verify(git => git.CreateBranch(It.Is<string>(s => s == vm.CurrentBranch), It.Is<string>(s => s == "bugBranch")));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIshiddenAfterSubmit()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                DisplayCreateBranchGrid = true
            };

            vm.CreateBranchOkButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayCreateBranchGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIshiddenAfterCancel()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                DisplayCreateBranchGrid = true
            };

            vm.CreateBranchCancelButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayCreateBranchGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchUserInputIsClearedAfterSubmit()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "newBranch"
            };

            vm.CreateBranchOkButtonCommand.Execute(null);

            Assert.AreEqual(string.Empty, vm.NewBranchName);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchUserInputIsClearedAfterCancel()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "newBranch"
            };

            vm.CreateBranchCancelButtonCommand.Execute(null);

            Assert.AreEqual(string.Empty, vm.NewBranchName);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsShownOnMergeClick()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            vm.MergeBranchCommand.Execute(null);

            Assert.IsTrue(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsNotShownWhenLocal_IsNull()
        {
            var vm = new BranchesViewViewModel();

            Assert.IsFalse(vm.MergeBranchCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewSelectedDestinationBranchIsCurrentBranch()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };
 
            Assert.AreEqual(_intialBranch.Name, vm.DestinationBranch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderMergesOnMergeViewSubmit()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                SourceBranch = "dev",
                DestinationBranch = "master"
            };

            vm.MergeBranchesOkButtonCommand.Execute(null);

            _provider.Verify(git => git.Merge("dev", "master"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsHiddenOnSuccessfulMerge()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                SourceBranch = "fizz",
                DestinationBranch = "buzz",
                DisplayMergeBranchesGrid = true
            };

            vm.MergeBranchesOkButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsHiddenOnCancel()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                DisplayMergeBranchesGrid = true
            };

            vm.MergeBranchesCancelButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ChangingSelectedBranchChecksOutThatBranch()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentBranch = "dev"
            };

            _provider.Verify(git => git.Checkout("dev"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void RefreshingViewShouldNotCheckoutBranch()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
            };

            vm.RefreshView();

            _provider.Verify(git => git.Checkout(It.IsAny<string>()), Times.Once);  //checkout when we first set provider
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnBranchChange_WhenCheckoutFails_ActionFailedEventIsRaised()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };
            var wasRaised = false;

            _provider.Setup(p => p.Checkout(It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.CurrentBranch = null;

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnMergeBranch_WhenCheckoutFails_ActionFailedEventIsRaised()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };
            var wasRaised = false;

            _provider.Setup(p => p.Merge(It.IsAny<string>(), It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.MergeBranchesOkButtonCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnDeleteBranch_WhenDeleteFails_ActionFailedEventIsRaised()
        {
            var branchName = "dev";

            var wasRaised = false;
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = branchName
            };

            _provider.Setup(p => p.DeleteBranch(It.Is<string>(b => b == branchName)))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.DeleteBranchToolbarButtonCommand.Execute(bool.TrueString);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnCreateBranch_WhenCreateFails_ActionFailedEventIsRaised()
        {
            var wasRaised = false;
            var branchName = "dev";

            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = branchName
            };

            _provider.Setup(p => p.CreateBranch(It.Is<string>(b => b == vm.CurrentBranch), It.Is<string>(b => b == branchName)))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.CreateBranchOkButtonCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void PublishPublishesBranch()
        {
            var branch = "dev";
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentUnpublishedBranch = branch
            };

            vm.PublishBranchToolbarButtonCommand.Execute(null);

            _provider.Verify(git => git.Publish(branch));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnpublishUnpublishesBranch()
        {
            var branch = "master";
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = branch
            };

            vm.UnpublishBranchToolbarButtonCommand.Execute(null);

            _provider.Verify(git => git.Unpublish(branch));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void PublishBranch_ActionFailedEventIsRaised()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };
            var wasRaised = false;

            _provider.Setup(p => p.Publish(It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.PublishBranchToolbarButtonCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnpublishBranch_ActionFailedEventIsRaised()
        {
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "master"
            };
            var wasRaised = false;

            _provider.Setup(p => p.Unpublish(It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.UnpublishBranchToolbarButtonCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
