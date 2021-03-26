using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;
using Times = Moq.Times;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    [Category("ComMock.ItByRef")]
    public class ItByRefTests
    {
        [Test]
        public void Basic_Ref_Setup()
        {
            var mock = new Mock<ITestRef>();
            var byRef = ItByRef<int>.Is(1, (ref int x) => x = 2);
            mock.Setup(x => x.DoInt(ref byRef.Value)).Callback(byRef.Callback);
            var obj = mock.Object;
            obj.DoInt(ref byRef.Value);

            Assert.AreEqual(2, byRef.Value);
            mock.Verify(x => x.DoInt(ref byRef.Value), Times.Once);
        }

        [Test]
        public void Multiple_Ref_Setup()
        {
            var mock = new Mock<ITestRef>();
            var byRef1 = ItByRef<int>.Is(1, (ref int x) => x = 2);
            var byRef2 = ItByRef<int>.Is(3, (ref int x) => x = 5);
            mock.Setup(x => x.DoInt(ref byRef1.Value)).Callback(byRef1.Callback);
            mock.Setup(x => x.DoInt(ref byRef2.Value)).Callback(byRef2.Callback);

            var obj = mock.Object;
            obj.DoInt(ref byRef1.Value);
            obj.DoInt(ref byRef2.Value);

            Assert.AreEqual(2, byRef1.Value);
            Assert.AreEqual(5, byRef2.Value);
            mock.Verify(x => x.DoInt(ref byRef1.Value), Times.Once);
            mock.Verify(x => x.DoInt(ref byRef2.Value), Times.Once);
        }

        [Test]
        public void Null_Ref_Setup()
        {
            var mock = new Mock<ITestRef>();
            var byRef = ItByRef<string>.Is(null, (ref string x) => x = string.Empty);
            mock.Setup(x => x.DoString(ref byRef.Value)).Callback(byRef.Callback);
            var obj = mock.Object;
            obj.DoString(ref byRef.Value);

            var testString = "abc";
            obj.DoString(ref testString);

            Assert.AreEqual(string.Empty, byRef.Value);
            Assert.AreEqual("abc", testString);

            mock.Verify(x => x.DoString(ref byRef.Value), Times.Once);
        }

        [Test]
        public void Basic_Ref_Setup_Returns()
        {
            var mock = new Mock<ITestRef>();
            var byRef = ItByRef<int>.Is(1);
            mock.Setup(x => x.ReturnInt(ref byRef.Value)).Returns(2);
            var obj = mock.Object;
            var actual = obj.ReturnInt(ref byRef.Value);

            var negativeRef = 0;
            var negativeActual = obj.ReturnInt(ref negativeRef);

            Assert.AreEqual(2, actual);
            Assert.AreEqual(0, negativeActual);
        }

        [Test]
        public void ItByRefMemberInfos_Is()
        {
            Assert.IsNotNull(ItByRefMemberInfos.Is(typeof(int)));
        }

        [Test]
        [TestCase(nameof(ITestRef.DoInt), 1)]
        [TestCase(nameof(ITestRef.DoString), "abc")]
        public void Test_ByRef_Setup(string memberName, object value)
        {
            var definitions = new SetupArgumentDefinitions
            {
                SetupArgumentDefinition.CreateIs(value)
            };
            var resolver = new SetupArgumentResolver();
            var builder = new SetupExpressionBuilder(typeof(ITestRef), new List<Type>(), resolver);
            
            var mock = new Mock<ITestRef>();
            var setupDatas = builder.CreateExpression(memberName, definitions);
            var setupData = setupDatas.First();

            var called = false;
            void Action()
            {
                called = true;
            }

            MockExpressionBuilder.Create(mock)
                .As(typeof(ITestRef))
                .Setup(setupData.SetupExpression, setupData.Args)
                .Callback(Action)
                .Execute();

            object refParam = null;
            switch (memberName)
            {
                case nameof(ITestRef.DoInt):
                    var refInt = (int) value;
                    mock.Object.DoInt(ref refInt);
                    refParam = refInt;
                    break;
                case nameof(ITestRef.DoString):
                    var refString = (string) value;
                    mock.Object.DoString(ref refString);
                    refParam = refString;
                    break;
                default:
                    Assert.Fail("Missing case for a member call");
                    return;
            }
            
            Assert.AreEqual(true, called);
            Assert.AreEqual(value, refParam);
        }
    }
}
