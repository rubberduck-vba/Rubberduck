using NUnit.Framework;
using Rubberduck.VBEditor.ComManagement;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ComSafeTests
    {
        [Test()]
        public void TryRemoveOnNewComSafe_ReturnsFalse()
        {
            var comSafe = new ComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test()]
        public void TryRemoveWithItemPreviouslyAddedToComSafe_ReturnsTrue()
        {
            var comSafe = new ComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsTrue(result);
        }

        [Test()]
        public void TryRemoveWithOtherItemPreviouslyAddedToComSafe_ReturnsFalse()
        {
            var comSafe = new ComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            var otherTestComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(otherTestComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test()]
        public void TryRemoveWithItemAndOtherItemPreviouslyAddedToComSafe_ReturnsTrue()
        {
            var comSafe = new ComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            var otherTestComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(otherTestComWrapper);
            comSafe.Add(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsTrue(result);
        }

        [Test()]
        public void SecondTryRemoveWithItemPreviouslyAddedToComSafe_ReturnsFalse()
        {
            var comSafe = new ComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test()]
        public void SecondTryRemoveWithItemPreviouslyAddedToComSafeTwice_ReturnsFalse()
        {
            var comSafe = new ComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test()]
        public void AddedSafeComWrapperGetsDisposedOnDisposalOfComSafe()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test()]
        public void SafeComWrapperAddedTwiceGetsDisposedOnceOnDisposalOfComSafe()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test()]
        public void RemovedSafeComWrapperDoesNotGetDisposedOnDisposalOfComSafe()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Never);
        }

        [Test()]
        public void SafeComWrapperRemovedAfterHavingBeenAddedTwiceDoesNotGetDisposedOnDisposalOfComSafe()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Never);
        }

        [Test()]
        public void AddedSafeComWrapperGetsDisposedOnDisposalOfAfterOtherItemGotRemovedComSafe()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            var otherTestComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(otherTestComWrapper);
            comSafe.TryRemove(otherTestComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test()]
        public void AddedSafeComWrapperDoesNotGetDisposedAgainOnSecondDisposalOfComSafe()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Dispose();
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test()]
        public void AfterDisposalTryRemoveReturnsFalseForAddedItem()
        {
            var comSafe = new ComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Dispose();
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }
    }
}