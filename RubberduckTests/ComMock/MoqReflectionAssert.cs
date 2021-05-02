using System;
using System.Linq;
using System.Linq.Expressions;
using Moq;
using Moq.Language;
using Moq.Language.Flow;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;

namespace RubberduckTests.ComMock
{
    internal delegate void Callback();

    [TestFixture]
    [Category("ComMocks.MoqReflection")]
    public class MoqReflectionAssert
    {
        [Test]
        public void As_Method_Exists()
        {
            var asMethod = MockMemberInfos.As(typeof(object));
            var foundMethod = typeof(Mock<>).GetMethods().Single(x => 
                x.Name == nameof(Mock<object>.As) &&
                x.IsGenericMethod && 
                x.GetGenericArguments().Length == 1 &&
                x.GetParameters().Length == 0);

            Assert.AreEqual(asMethod.GetType(), foundMethod.GetType());
        }

        [Test]
        public void Setup_Method_Exists()
        {
            var setupMethod = MockMemberInfos.Setup(typeof(Mock<object>), null);
            var foundMethod = typeof(Mock<>).GetMethods().Single(x =>
                x.Name == nameof(Mock<object>.Setup) &&
                x.IsGenericMethod &&
                x.GetGenericArguments().Length == 1 &&
                x.GetParameters().Length == 1);

            Assert.AreEqual(setupMethod.GetType(), foundMethod.GetType());
        }

        [Test]
        public void Setup_Without_Returns_Is_Executed_On_ITest1()
        {
            var mocked = new Mock<ITest1>();
            Expression<Action<ITest1>> expression = x => x.DoThis();
            var setupMethod = MockMemberInfos.Setup(mocked.GetType(), null);

            // We need to verify this succeeds
            setupMethod.Invoke(mocked, new object[] { expression });
        }

        [Test]
        public void Setup_Without_Returns_Is_Executed_On_ITest2()
        {
            var mocked = new Mock<ITest2>();
            Expression<Action<ITest2>> expression = x => x.DoThat();
            var setupMethod = MockMemberInfos.Setup(mocked.GetType(), null);

            // We need to verify this succeeds
            setupMethod.Invoke(mocked, new object[] { expression });
        }

        [Test]
        public void Setup_With_Returns_Is_Executed_On_ITest1()
        {
            var mocked = new Mock<ITest1>();
            Expression<Func<ITest1, int>> expression = x => x.DoThis();
            var setupMethod = MockMemberInfos.Setup(mocked.GetType(), typeof(int));

            // We need to verify this succeeds
            setupMethod.Invoke(mocked, new object[] {expression});
        }

        [Test]
        public void Setup_With_Returns_Is_Executed_On_ITest2()
        {
            var mocked = new Mock<ITest2>();
            Expression<Func<ITest2, string>> expression = x => x.DoThat();
            var setupMethod = MockMemberInfos.Setup(mocked.GetType(), typeof(string));

            // We need to verify this succeeds
            setupMethod.Invoke(mocked, new object[] { expression });
        }

        [Test]
        public void Returns_Method_Exists()
        {
            var mocked = new Mock<ITest1>();
            var setup = mocked.Setup(x => x.DoThis());
            var returnMethod = MockMemberInfos.Returns(setup.GetType());
            var foundMethod = typeof(IReturns<,>).GetMethods().Single(x =>
                x.Name == nameof(IReturns<object, object>.Returns) &&
                x.IsGenericMethod &&
                x.GetGenericArguments().Length == 1 &&
                x.GetParameters().Length == 1);

            Assert.AreEqual(returnMethod.GetType(), foundMethod.GetType());
        }

        [Test]
        public void Returns_Is_Executed_On_ITest1()
        {
            const int expected = 42;
            var mocked = new Mock<ITest1>();
            var setup = mocked.Setup(x => x.DoThis());
            var returnMethod = MockMemberInfos.Returns(setup.GetType());

            returnMethod.Invoke(setup, new object[] { expected });

            var test = mocked.Object;
            Assert.AreEqual(expected, test.DoThis());
        }

        [Test]
        public void Returns_Is_Executed_On_ITest2()
        {
            const string expected = "abc";
            var mocked = new Mock<ITest2>();
            var setup = mocked.Setup(x => x.DoThat());
            var returnMethod = MockMemberInfos.Returns(setup.GetType());

            returnMethod.Invoke(setup, new object[] { expected });

            var test = mocked.Object;
            Assert.AreEqual(expected, test.DoThat());
        }

        [Test]
        public void Callback_Method_Exists()
        {
            var mocked = new Mock<ITest1>();
            var setup = mocked.Setup(x => x.DoThis());
            var returnMethod = MockMemberInfos.Callback(setup.GetType());
            var foundMethod = typeof(ICallback).GetMethods().Single(x =>
                x.Name == nameof(ISetup<object>.Callback) &&
                !x.IsGenericMethod &&
                x.GetGenericArguments().Length == 0 &&
                x.GetParameters().Length == 1 &&
                x.GetParameters()[0].ParameterType == typeof(Delegate));

            Assert.AreEqual(returnMethod.GetType(), foundMethod.GetType());
        }

        [Test]
        public void Callback_Is_Executed_On_ITest1()
        {
            const bool expected = true;
            var actual = false;
            void Expression() => actual = true;

            var mocked = new Mock<ITest1>();
            var setup = mocked.Setup(x => x.DoThis());
            var callbackMethod = MockMemberInfos.Callback(setup.GetType());

            callbackMethod.Invoke(setup, new object[] { (Callback)Expression });

            var test = mocked.Object;
            test.DoThis();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Callback_Is_Executed_On_ITest2()
        {
            const bool expected = true;
            var actual = false;
            void Expression() => actual = true;

            var mocked = new Mock<ITest2>();
            var setup = mocked.Setup(x => x.DoThat());
            var callbackMethod = MockMemberInfos.Callback(setup.GetType());

            callbackMethod.Invoke(setup, new object[] { (Callback) Expression });

            var test = mocked.Object;
            test.DoThat();
            Assert.AreEqual(expected, actual);
        }
    }
}
