using NUnit.Framework;
using Rubberduck.VBEditor.ComManagement;
using System;
using Moq;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ReferenceEqualityComparerTests
    {
        [Test]
        [Category("COM")]
        public void AnObjectIsEqualToItself()
        {
            var referenceComparer = new ReferenceEqualityComparer();

            var obj1 = new object();
            var obj2 = obj1;

            var consideredEqual = referenceComparer.Equals(obj1, obj2);

            Assert.IsTrue(consideredEqual);
        }

        [Test]
        [Category("COM")]
        public void DifferentObjectsEqualAsIEquatableAreNotConsideredEqual()
        {
            var referenceComparer = new ReferenceEqualityComparer();

            var mock1 = new Mock<IEquatable<object>>();
            mock1.Setup(obj => obj.Equals(It.IsAny<object>())).Returns(true);
            var mock2 = new Mock<IEquatable<object>>();
            mock2.Setup(obj => obj.Equals(It.IsAny<object>())).Returns(true);
            var obj1 = mock1.Object;
            var obj2 = mock2.Object;

            var consideredEqual = referenceComparer.Equals(obj1, obj2);

            Assert.IsFalse(consideredEqual);
        }

        [Test]
        [Category("COM")]
        public void GetHashCodeIgnoresHashAsIEquatable()
        {
            var referenceComparer = new ReferenceEqualityComparer();

            var mock = new Mock<IEquatable<object>>();
            mock.Setup(obj => obj.GetHashCode()).Returns(0);
            var objct = mock.Object;

            var hashCode = referenceComparer.GetHashCode(objct);

            Assert.AreNotEqual(0, hashCode);
        }
    }
}