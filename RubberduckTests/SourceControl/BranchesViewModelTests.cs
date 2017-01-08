﻿using System.Collections.Generic;
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
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.SetupGet(git => git.CurrentBranch).Returns(_branches[1]);

            //act
            vm.RefreshView();

            //assert
            Assert.AreEqual(_provider.Object.CurrentBranch.Name, vm.CurrentBranch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void PublishedBranchesAreListed()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //assert
            var expected = new List<string> { "master" };
            CollectionAssert.AreEqual(expected, vm.PublishedBranches.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnPublishedBranchesAreListed()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //assert
            var expected = new List<string> { "dev" };
            CollectionAssert.AreEqual(expected, vm.UnpublishedBranches.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnlyLocalBranchesInBranches()
        {
            //arrange 
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //assert
            var expected = new List<string> { "master", "dev" };
            CollectionAssert.AreEquivalent(expected, vm.LocalBranches.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void HeadIsNotIncludedInPublishedBranches()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //assert
            CollectionAssert.DoesNotContain(vm.PublishedBranches.ToList(), "HEAD");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void DeleteBranchDisabled_BranchIsActive()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "master"
            };

            //Assert
            Assert.IsFalse(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.TrueString));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void DeleteBranchEnabled_BranchIsNotActive()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = "bugbranch"
            };

            //Assert
            Assert.IsTrue(vm.DeleteBranchToolbarButtonCommand.CanExecute(bool.TrueString));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void DeleteBranch_BranchIsNotActive_BranchIsRemoved()
        {
            //arrange
            var firstBranchName = "master";
            var secondBranchName = "bugBranch";

            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //Assert
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

            //act
            vm.DeleteBranchToolbarButtonCommand.Execute(bool.FalseString);

            //Assert
            _provider.Verify(p => p.DeleteBranch(secondBranchName));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIsShownOnCreateBranch()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //act
            vm.NewBranchCommand.Execute(null);

            //Assert
            Assert.IsTrue(vm.DisplayCreateBranchGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeBranchViewIsShownOnCreateBranch()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //act
            vm.NewBranchCommand.Execute(null);

            //Assert
            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_BranchExists()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute("master"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_ValidBranchName()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch"
            };

            //Assert
            Assert.IsTrue(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsSpace()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsTwoConsecutiveDots()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug..branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsTilde()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug~branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsCaret()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug^branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsColon()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug:branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsQuestionMark()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug?branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsAsteriks()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug*branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsOpenBracket()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug[branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsTwoConsecutiveSlashes()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug//branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameStartsWithSlash()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "/bugBranch"
            };

            //Assert
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

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameEndsWithDot()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch."
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameIsAtSign()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "@"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsConsecutiveAtSignAndOpenBrace()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug@{branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsBackslash()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug\\branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsSlashSectionStartingWithDot()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug/.branch"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranch_NameContainsSlashSectionEndingWithDotlock()
        {
            //arrange
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bug/branch.lock"
            };

            //Assert
            Assert.IsFalse(vm.CreateBranchOkButtonCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIsNotShownWhenLocal_IsNull()
        {
            //arrange
            var vm = new BranchesViewViewModel();

            //Assert
            Assert.IsFalse(vm.NewBranchCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderCallsCreateBranchOnCreateBranchConfirm()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "bugBranch"
            };

            //act
            vm.CreateBranchOkButtonCommand.Execute(null);

            //assert
            _provider.Verify(git => git.CreateBranch(It.Is<string>(s => s == vm.CurrentBranch), It.Is<string>(s => s == "bugBranch")));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIshiddenAfterSubmit()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                DisplayCreateBranchGrid = true
            };

            //act
            vm.CreateBranchOkButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(vm.DisplayCreateBranchGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchViewIshiddenAfterCancel()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                DisplayCreateBranchGrid = true
            };

            //act
            vm.CreateBranchCancelButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(vm.DisplayCreateBranchGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchUserInputIsClearedAfterSubmit()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "newBranch"
            };

            //act
            vm.CreateBranchOkButtonCommand.Execute(null);

            //assert
            Assert.AreEqual(string.Empty, vm.NewBranchName);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CreateBranchUserInputIsClearedAfterCancel()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                NewBranchName = "newBranch"
            };

            //act
            vm.CreateBranchCancelButtonCommand.Execute(null);

            //assert
            Assert.AreEqual(string.Empty, vm.NewBranchName);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsShownOnMergeClick()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //act
            vm.MergeBranchCommand.Execute(null);

            //Assert
            Assert.IsTrue(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsNotShownWhenLocal_IsNull()
        {
            //arrange
            var vm = new BranchesViewViewModel();

            //Assert
            Assert.IsFalse(vm.MergeBranchCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewSelectedDestinationBranchIsCurrentBranch()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object
            };

            //assert 
            Assert.AreEqual(_intialBranch.Name, vm.DestinationBranch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderMergesOnMergeViewSubmit()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                SourceBranch = "dev",
                DestinationBranch = "master"
            };

            //act
            vm.MergeBranchesOkButtonCommand.Execute(null);

            //assert
            _provider.Verify(git => git.Merge("dev", "master"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsHiddenOnSuccessfulMerge()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                SourceBranch = "fizz",
                DestinationBranch = "buzz",
                DisplayMergeBranchesGrid = true
            };

            //act
            vm.MergeBranchesOkButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void MergeViewIsHiddenOnCancel()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                DisplayMergeBranchesGrid = true
            };

            //act
            vm.MergeBranchesCancelButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(vm.DisplayMergeBranchesGrid);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ChangingSelectedBranchChecksOutThatBranch()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentBranch = "dev"
            };

            //assert
            _provider.Verify(git => git.Checkout("dev"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void RefreshingViewShouldNotCheckoutBranch()
        {
            //arrange
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
            };

            //act
            vm.RefreshView();

            //assert
            _provider.Verify(git => git.Checkout(It.IsAny<string>()), Times.Once);  //checkout when we first set provider
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnBranchChange_WhenCheckoutFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.CurrentBranch = null;

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnMergeBranch_WhenCheckoutFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.MergeBranchesOkButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnDeleteBranch_WhenDeleteFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.DeleteBranchToolbarButtonCommand.Execute(bool.TrueString);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnCreateBranch_WhenCreateFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.CreateBranchOkButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void PublishPublishesBranch()
        {
            //arrange
            var branch = "dev";
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentUnpublishedBranch = branch
            };

            //act
            vm.PublishBranchToolbarButtonCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Publish(branch));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnpublishUnpublishesBranch()
        {
            //arrange
            var branch = "master";
            var vm = new BranchesViewViewModel
            {
                Provider = _provider.Object,
                CurrentPublishedBranch = branch
            };

            //act
            vm.UnpublishBranchToolbarButtonCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Unpublish(branch));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void PublishBranch_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.PublishBranchToolbarButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnpublishBranch_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.UnpublishBranchToolbarButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
