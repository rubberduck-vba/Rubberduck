using NUnit.Framework;
using Rubberduck.VBEditor.ComManagement;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public abstract class ComSafeTestBase
    {
        protected abstract IComSafe TestComSafe();

        [Test]
        [Category("COM")]
        public void TryRemoveOnNewComSafe_ReturnsFalse()
        {
            var comSafe = TestComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test]
        [Category("COM")]
        public void TryRemoveWithItemPreviouslyAddedToComSafe_ReturnsTrue()
        {
            var comSafe = TestComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsTrue(result);
        }

        [Test]
        [Category("COM")]
        public void TryRemoveWithOtherItemPreviouslyAddedToComSafe_ReturnsFalse()
        {
            var comSafe = TestComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            var otherTestComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(otherTestComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test]
        [Category("COM")]
        public void TryRemoveWithItemAndOtherItemPreviouslyAddedToComSafe_ReturnsTrue()
        {
            var comSafe = TestComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            var otherTestComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(otherTestComWrapper);
            comSafe.Add(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsTrue(result);
        }

        [Test]
        [Category("COM")]
        public void SecondTryRemoveWithItemPreviouslyAddedToComSafe_ReturnsFalse()
        {
            var comSafe = TestComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test]
        [Category("COM")]
        public void SecondTryRemoveWithItemPreviouslyAddedToComSafeTwice_ReturnsFalse()
        {
            var comSafe = TestComSafe();
            var testComWrapper = new Mock<ISafeComWrapper>().Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [Test]
        [Category("COM")]
        public void AddedSafeComWrapperGetsDisposedOnDisposalOfComSafe()
        {
            var comSafe = TestComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void SafeComWrapperAddedTwiceGetsDisposedOnceOnDisposalOfComSafe()
        {
            var comSafe = TestComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void RemovedSafeComWrapperDoesNotGetDisposedOnDisposalOfComSafe()
        {
            var comSafe = TestComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void SafeComWrapperRemovedAfterHavingBeenAddedTwiceDoesNotGetDisposedOnDisposalOfComSafe()
        {
            var comSafe = TestComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Add(testComWrapper);
            comSafe.TryRemove(testComWrapper);
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void AddedSafeComWrapperGetsDisposedOnDisposalOfAfterOtherItemGotRemovedComSafe()
        {
            var comSafe = TestComSafe();
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

        [Test]
        [Category("COM")]
        public void AddedSafeComWrapperDoesNotGetDisposedAgainOnSecondDisposalOfComSafe()
        {
            var comSafe = TestComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Dispose();
            comSafe.Dispose();

            mock.Verify(wrapper => wrapper.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void AfterDisposalTryRemoveReturnsFalseForAddedItem()
        {
            var comSafe = TestComSafe();
            var mock = new Mock<ISafeComWrapper>();
            mock.Setup(wrapper => wrapper.Dispose());

            var testComWrapper = mock.Object;
            comSafe.Add(testComWrapper);
            comSafe.Dispose();
            var result = comSafe.TryRemove(testComWrapper);

            Assert.IsFalse(result);
        }

        [TestFixture()]
        public class StrongComSafeTests : ComSafeTestBase
        {
            protected override IComSafe TestComSafe()
            {
                return new StrongComSafe();
            }
        }

        [TestFixture()]
        public class WeakComSafeTests : ComSafeTestBase
        {
            protected override IComSafe TestComSafe()
            {
                return new WeakComSafe();
            }
        }
    }
}