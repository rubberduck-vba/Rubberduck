using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestFixture]
    public class BranchesViewModelTests
    {
        private Mock<ISourceControlProvider> _provider;
        private Branch _intialBranch;
        private List<IBranch> _branches;

        [Category("SourceControl")]
        [SetUp]
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

        [Category("SourceControl")]
        [Test]
        public void SelectedBranchShouldBeCurrentBranchAfterRefresh()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            _provider.SetupGet(git => git.CurrentBranch).Returns(_branches[1]);

            vm.RefreshView();

            Assert.AreEqual(_provider.Object.CurrentBranch.Name, vm.CurrentBranch);
        }

        [Category("SourceControl")]
        [Test]
        public void PublishedBranchesAreListed()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            var expected = new List<string> { "master" };
            CollectionAssert.AreEqual(expected, vm.PublishedBranches.ToList());
        }

        [Category("SourceControl")]
        [Test]
        public void UnPublishedBranchesAreListed()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            var expected = new List<string> { "dev" };
            CollectionAssert.AreEqual(expected, vm.UnpublishedBranches.ToList());
        }

        [Category("SourceControl")]
        [Test]
        public void OnlyLocalBranchesInBranches()
        { 
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            var expected = new List<string> { "master", "dev" };
            CollectionAssert.AreEquivalent(expected, vm.LocalBranches.ToList());
        }

        [Category("SourceControl")]
        [Test]
        public void HeadIsNotIncludedInPublishedBranches()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            CollectionAssert.DoesNotContain(vm.PublishedBranches.ToList(), "HEAD");
        }

        [Category("SourceControl")]
        [Test]
        public void DeleteBranchDisabled_BranchIsActive()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "master"
            };

            Assert.IsFalse(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.TrueString));
        }

        [Category("SourceControl")]
        [Test]
        public void DeleteBranchEnabled_BranchIsNotActive()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "bugbranch"
            };

            Assert.IsTrue(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.TrueString));
        }

        [Category("SourceControl")]
        [Test]
        public void DeleteBranch_BranchIsNotActive_BranchIsRemoved()
        {
            var firstBranchName = "master";
            var secondBranchName = "bugBranch";

            var vm = new BranchesPanelViewModel
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

        [Category("SourceControl")]
        [Test]
        public void CreateBranchViewIsShownOnCreateBranch()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            vm.NewBranchCommand.Execute(null);

            Assert.IsTrue(vm.DisplayCreateBranchGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void MergeBranchViewIsShownOnCreateBranch()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            vm.NewBranchCommand.Execute(null);

            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_BranchExists()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute("master"));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_ValidBranchName()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch"
            };

            Assert.IsTrue(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsSpace()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsTwoConsecutiveDots()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug..branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsTilde()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug~branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsCaret()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug^branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsColon()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug:branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsQuestionMark()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug?branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsAsteriks()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug*branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsOpenBracket()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug[branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsTwoConsecutiveSlashes()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug//branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameStartsWithSlash()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "/bugBranch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameEndsWithSlash()
        {            //arrange
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch/"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameEndsWithDot()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch."
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameIsAtSign()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "@"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsConsecutiveAtSignAndOpenBrace()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug@{branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsBackslash()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug\\branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsSlashSectionStartingWithDot()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug/.branch"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranch_NameContainsSlashSectionEndingWithDotlock()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug/branch.lock"
            };

            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranchViewIsNotShownWhenLocal_IsNull()
        {
            var vm = new BranchesPanelViewModel();

            Assert.IsFalse(vm.NewBranchCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void ProviderCallsCreateBranchOnCreateBranchConfirm()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch"
            };

            vm.CreateBranchOkButtonCommand.Execute(null);

            _provider.Verify(git => git.CreateBranch(It.Is<string>(s => s == vm.CurrentBranch), It.Is<string>(s => s == "bugBranch")));
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranchViewIshiddenAfterSubmit()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                DisplayCreateBranchGrid = true
            };

            vm.CreateBranchOkButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayCreateBranchGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranchViewIshiddenAfterCancel()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                DisplayCreateBranchGrid = true
            };

            vm.CreateBranchCancelButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayCreateBranchGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranchUserInputIsClearedAfterSubmit()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "newBranch"
            };

            vm.CreateBranchOkButtonCommand.Execute(null);

            Assert.AreEqual(string.Empty, vm.NewBranchName);
        }

        [Category("SourceControl")]
        [Test]
        public void CreateBranchUserInputIsClearedAfterCancel()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "newBranch"
            };

            vm.CreateBranchCancelButtonCommand.Execute(null);

            Assert.AreEqual(string.Empty, vm.NewBranchName);
        }

        [Category("SourceControl")]
        [Test]
        public void MergeViewIsShownOnMergeClick()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };

            vm.MergeBranchCommand.Execute(null);

            Assert.IsTrue(vm.DisplayMergeBranchesGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void MergeViewIsNotShownWhenLocal_IsNull()
        {
            var vm = new BranchesPanelViewModel();

            Assert.IsFalse(vm.MergeBranchCommand.CanExecute(null));
        }

        [Category("SourceControl")]
        [Test]
        public void MergeViewSelectedDestinationBranchIsCurrentBranch()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object
            };
 
            Assert.AreEqual(_intialBranch.Name, vm.DestinationBranch);
        }

        [Category("SourceControl")]
        [Test]
        public void ProviderMergesOnMergeViewSubmit()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                SourceBranch = "dev",
                DestinationBranch = "master"
            };

            vm.MergeBranchesOkButtonCommand.Execute(null);

            _provider.Verify(git => git.Merge("dev", "master"));
        }

        [Category("SourceControl")]
        [Test]
        public void MergeViewIsHiddenOnSuccessfulMerge()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                SourceBranch = "fizz",
                DestinationBranch = "buzz",
                DisplayMergeBranchesGrid = true
            };

            vm.MergeBranchesOkButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void MergeViewIsHiddenOnCancel()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                DisplayMergeBranchesGrid = true
            };

            vm.MergeBranchesCancelButtonCommand.Execute(null);

            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [Category("SourceControl")]
        [Test]
        public void ChangingSelectedBranchChecksOutThatBranch()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                CurrentBranch = "dev"
            };

            _provider.Verify(git => git.Checkout("dev"));
        }

        [Category("SourceControl")]
        [Test]
        public void RefreshingViewShouldNotCheckoutBranch()
        {
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
            };

            vm.RefreshView();

            _provider.Verify(git => git.Checkout(It.IsAny<string>()), Times.Once);  //checkout when we first set provider
        }

        [Category("SourceControl")]
        [Test]
        public void OnBranchChange_WhenCheckoutFails_ActionFailedEventIsRaised()
        {
            var vm = new BranchesPanelViewModel
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

        [Category("SourceControl")]
        [Test]
        public void OnMergeBranch_WhenCheckoutFails_ActionFailedEventIsRaised()
        {
            var vm = new BranchesPanelViewModel
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

        [Category("SourceControl")]
        [Test]
        public void OnDeleteBranch_WhenDeleteFails_ActionFailedEventIsRaised()
        {
            var branchName = "dev";

            var wasRaised = false;
            var vm = new BranchesPanelViewModel
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

        [Category("SourceControl")]
        [Test]
        public void OnCreateBranch_WhenCreateFails_ActionFailedEventIsRaised()
        {
            var wasRaised = false;
            var branchName = "dev";

            var vm = new BranchesPanelViewModel
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

        [Category("SourceControl")]
        [Test]
        public void PublishPublishesBranch()
        {
            var branch = "dev";
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                CurrentUnpublishedBranch = branch
            };

            vm.PublishBranchToolbarButtonCommand.Execute(null);

            _provider.Verify(git => git.Publish(branch));
        }

        [Category("SourceControl")]
        [Test]
        public void UnpublishUnpublishesBranch()
        {
            var branch = "master";
            var vm = new BranchesPanelViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = branch
            };

            vm.UnpublishBranchToolbarButtonCommand.Execute(null);

            _provider.Verify(git => git.Unpublish(branch));
        }

        [Category("SourceControl")]
        [Test]
        public void PublishBranch_ActionFailedEventIsRaised()
        {
            var vm = new BranchesPanelViewModel
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

        [Category("SourceControl")]
        [Test]
        public void UnpublishBranch_ActionFailedEventIsRaised()
        {
            var vm = new BranchesPanelViewModel
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
