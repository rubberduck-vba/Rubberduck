using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    public class EngineTests
    {
        [Test]
        [Category("UnitTesting")]
        public void TestEngine_ExposesTestMethod_AndRaisesRefresh()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var typeLibApi = new Mock<IVBETypeLibsAPI>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var fakesFactory = new Mock<IFakesFactory>();
            var dispatcher = new Mock<IUiDispatcher>();

            using (var state = parser.State)
            {
                var engine = new TestEngine(state, fakesFactory.Object, typeLibApi.Object, wrapperProvider.Object, dispatcher.Object);
                int refreshes = 0;
                engine.TestsRefreshed += (sender, args) => refreshes++;
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun())
                {
                    Assert.Inconclusive("Parser Error");
                }

                Assert.AreEqual(1, engine.Tests.Count());
                Assert.AreEqual(1, refreshes);
            }
        }

        [Test]
        [Category("UnitTesting")]
        public void TestEngine_RaisesRefreshEvent_EveryParserRun()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var typeLibApi = new Mock<IVBETypeLibsAPI>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var fakesFactory = new Mock<IFakesFactory>();
            var dispatcher = new Mock<IUiDispatcher>();

            using (var state = parser.State)
            {
                var engine = new TestEngine(state, fakesFactory.Object, typeLibApi.Object, wrapperProvider.Object, dispatcher.Object);
                const int parserRuns = 5;
                int refreshes = 0;
                engine.TestsRefreshed += (sender, args) => refreshes++;
                for (int i = 0; i < parserRuns; i++)
                {
                    parser.Parse(new CancellationTokenSource());
                }
                if (!engine.CanRun())
                {
                    Assert.Inconclusive("Parser Error");
                }

                Assert.AreEqual(1, engine.Tests.Count());
                Assert.AreEqual(parserRuns, refreshes);
            }
        }

        [Test]
        [Category("UnitTesting")]
        public void TestEngine_Run_RaisesCompletionEvent_Success()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var typeLibApi = new Mock<IVBETypeLibsAPI>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var typeLibMock = new Mock<ITypeLibWrapper>();
            wrapperProvider.Setup(p => p.TypeLibWrapperFromProject(It.IsAny<string>()))
                            .Returns(typeLibMock.Object)
                            .Verifiable();
            typeLibApi.Setup(api => api.ExecuteCode(typeLibMock.Object, It.IsAny<string>(), It.IsAny<string>(), It.Is<object[]>(v => v == null)))
                .Verifiable();

            var fakesFactory = new Mock<IFakesFactory>();
            var createdFakes = new Mock<IFakes>();
            fakesFactory.Setup(factory => factory.Create())
                .Returns(createdFakes.Object);

            var dispatcher = new Mock<IUiDispatcher>();
            dispatcher.Setup(d => d.InvokeAsync(It.IsAny<Action>()))
                      .Callback((Action action) => action.Invoke())
                      .Verifiable();

            var completionEvents = new List<TestCompletedEventArgs>();
            using (var state = parser.State)
            {
                var engine = new TestEngine(state, fakesFactory.Object, typeLibApi.Object, wrapperProvider.Object, dispatcher.Object);
                engine.TestCompleted += (source, args) => completionEvents.Add(args);
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun())
                {
                    Assert.Inconclusive("Parser Error");
                }
                engine.Run(engine.Tests);
            }
            Mock.Verify(dispatcher, typeLibApi, wrapperProvider);
            Assert.AreEqual(1, completionEvents.Count);
            Assert.AreEqual(new TestResult(TestOutcome.Succeeded), completionEvents.First().Result);
        }

        //[Test]
        //public void TestEngine_SuccessfulTests()
        //{
        //    var actual = _engine.PassedTests().First();

        //    Assert.AreEqual(_successfulMethod, actual);
        //}

        //[Test]
        //public void TestEngine_NotRunTests()
        //{
        //    var actual = _engine.NotRunTests().First();

        //    Assert.AreEqual(_notRunMethod, actual);
        //}

        //[Test]
        //public void TestEngine_LastRunTests_ReturnsAllRunTests()
        //{
        //    var actual = _engine.LastRunTests().ToList();
        //    var expected = new List<TestMethod>()
        //    {
        //        _failedMethod, _inconclusiveMethod, _successfulMethod
        //    };

        //    CollectionAssert.AreEquivalent(expected, actual);
        //}

        //[Test]
        //public void TestEngine_LastRunTests_Successful()
        //{
        //    var actual = _engine.LastRunTests(TestOutcome.Succeeded).First();

        //    Assert.AreEqual(_successfulMethod, actual);
        //}

        //[Test]
        //public void TestEngine_LastRunTests_Failed()
        //{
        //    var actual = _engine.LastRunTests(TestOutcome.Failed).First();

        //    Assert.AreEqual(_failedMethod, actual);
        //}

        //[Test]
        //public void TestEngine_LastRunTests_Inconclusive()
        //{
        //    var actual = _engine.LastRunTests(TestOutcome.Inconclusive).First();

        //    Assert.AreEqual(_inconclusiveMethod, actual);
        //}

        //[Test]
        //public void TestEngine_Run_ModuleIntialize_IsRunOnce()
        //{
        //    //arrange
        //    _engine.ModuleInitialize += CatchEvent;

        //    var tests = _engine.AllTests.Keys;

        //    //act
        //    _engine.Run(tests);

        //    Assert.IsTrue(_wasEventRaised, "Module Intialize was not run.");
        //    Assert.AreEqual(1, _eventCount, "Module Intialzie expected to be run once.");
        //}

        //[Test]
        //public void TestEngine_Run_ModuleCleanup_IsRunOnce()
        //{
        //    //arrange
        //    _engine.ModuleCleanup += CatchEvent;

        //    //act
        //    _engine.Run(_engine.AllTests.Keys);

        //    //assert
        //    Assert.IsTrue(_wasEventRaised, "Module Cleanup was not run.");
        //    Assert.AreEqual(1, _eventCount, "Module Cleanup expected to be run once.");
        //}

        //[Test]
        //public void TestEngine_Run_MethodIntialize_IsRunForEachTestMethod()
        //{
        //    //arrange
        //    var expectedCount = _engine.AllTests.Count;
        //    _engine.MethodInitialize += CatchEvent;

        //    //act
        //    _engine.Run(_engine.AllTests.Keys);

        //    //assert
        //    Assert.IsTrue(_wasEventRaised, "Method Intialize was not run.");
        //    Assert.AreEqual(expectedCount, _eventCount, "Method Intialized was expected to be run {0} times", expectedCount);
        //}

        //[Test]
        //public void TestEngine_Run_MethodCleanup_IsRunForEachTestMethod()
        //{
        //    //arrange
        //    var expectedCount = _engine.AllTests.Count;
        //    _engine.MethodCleanup += CatchEvent;

        //    //act
        //    _engine.Run(_engine.AllTests.Keys);

        //    //assert
        //    Assert.IsTrue(_wasEventRaised, "Method Initialize was not run.");
        //    Assert.AreEqual(expectedCount, _eventCount, "Method Initialized was expected to be run {0} times", expectedCount);
        //}

        //[Test]
        //public void TestEngine_Run_TestCompleteIsRaisedForEachTestMethod()
        //{
        //    //arrange
        //    var expectedCount = _engine.AllTests.Count;
        //    _engine.TestCompleted += EngineOnTestComplete;

        //    //act
        //    _engine.Run(_engine.AllTests.Keys);

        //    //assert
        //    Assert.IsTrue(_wasEventRaised, "TestCompleted event was not raised.");
        //    Assert.AreEqual(expectedCount, _eventCount, "TestCompleted event was expected to be raised {0} times.", expectedCount);
        //}

        //[Test]
        //public void TestEngine_Run_WhenTestListIsEmpty_Bail()
        //{
        //    //arrange 
        //    _engine.MethodInitialize += CatchEvent;

        //    //act
        //    _engine.Run(new List<TestMethod>());

        //    //assert
        //    Assert.IsFalse(_wasEventRaised, "No methods should run when passed an empty list of tests.");
        //}

        //private void EngineOnTestComplete(object sender, TestCompletedEventArgs testCompletedEventArgs)
        //{
        //    CatchEvent();
        //}

        //private void CatchEvent(object sender, TestModuleEventArgs e)
        //{
        //    CatchEvent();
        //}

        //private void CatchEvent()
        //{
        //    _wasEventRaised = true;
        //    _eventCount++;
        //}


        private const string RawInput = @"Option Explicit
Option Private Module

{0}
Private Assert As New Rubberduck.AssertClass

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
";


        private string GetTestModuleInput
        {
            get { return string.Format(RawInput, "'@TestModule"); }
        }
    }
}
