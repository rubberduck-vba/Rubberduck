using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers;
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
        [Category("Unit Testing")]
        public void TestEngine_ExposesTestMethod_AndRaisesRefresh()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var interaction = new Mock<IVBEInteraction>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var fakesFactory = new Mock<IFakesFactory>();
            var dispatcher = new Mock<IUiDispatcher>();
            dispatcher.Setup(d => d.InvokeAsync(It.IsAny<Action>()))
              .Callback((Action action) => action.Invoke())
              .Verifiable();

            using (var state = parser.State)
            {
                var engine = new TestEngine(state, fakesFactory.Object, interaction.Object, wrapperProvider.Object, dispatcher.Object, vbe);
                int refreshes = 0;
                engine.TestsRefreshed += (sender, args) => refreshes++;
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun)
                {
                    Assert.Inconclusive("Parser Error");
                }

                Assert.AreEqual(1, engine.Tests.Count());
                Assert.AreEqual(1, refreshes);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_RaisesRefreshEvent_EveryParserRun()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var interaction = new Mock<IVBEInteraction>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var fakesFactory = new Mock<IFakesFactory>();
            var dispatcher = new Mock<IUiDispatcher>();
            dispatcher.Setup(d => d.InvokeAsync(It.IsAny<Action>()))
              .Callback((Action action) => action.Invoke())
              .Verifiable();

            using (var state = parser.State)
            {
                var engine = new TestEngine(state, fakesFactory.Object, interaction.Object, wrapperProvider.Object, dispatcher.Object, vbe);
                const int parserRuns = 5;
                int refreshes = 0;
                engine.TestsRefreshed += (sender, args) => refreshes++;
                for (int i = 0; i < parserRuns; i++)
                {
                    parser.Parse(new CancellationTokenSource());
                }
                if (!engine.CanRun)
                {
                    Assert.Inconclusive("Parser Error");
                }

                Assert.AreEqual(1, engine.Tests.Count());
                Assert.AreEqual(parserRuns, refreshes);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Run_RaisesCompletionEvent_Success()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var interaction = new Mock<IVBEInteraction>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var typeLibMock = new Mock<ITypeLibWrapper>();
            wrapperProvider.Setup(p => p.TypeLibWrapperFromProject(It.IsAny<string>()))
                            .Returns(typeLibMock.Object)
                            .Verifiable();
            long durationStub;
            interaction.Setup(ia => ia.RunTestMethod(typeLibMock.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out durationStub))
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
                var engine = new TestEngine(state, fakesFactory.Object, interaction.Object, wrapperProvider.Object, dispatcher.Object, vbe);
                engine.TestCompleted += (source, args) => completionEvents.Add(args);
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun)
                {
                    Assert.Inconclusive("Parser Error");
                }
                engine.Run(engine.Tests);
            }
            Mock.Verify(dispatcher, interaction, wrapperProvider);
            Assert.AreEqual(1, completionEvents.Count);
            Assert.AreEqual(new TestResult(TestOutcome.Succeeded), completionEvents.First().Result);
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Run_AndAssertSuccess_RaisesCompletionEvent_Success()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var interaction = new Mock<IVBEInteraction>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var typeLibMock = new Mock<ITypeLibWrapper>();

            long durationStub;
            interaction.Setup(ia => ia.RunTestMethod(typeLibMock.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out durationStub))
                .Callback(new RunTestMethodCallback((ITypeLibWrapper _, TestMethod method, EventHandler<AssertCompletedEventArgs> assertHandler, out long duration) =>
                {
                    duration = 0;
                    AssertHandler.OnAssertCompleted += assertHandler;
                    AssertHandler.OnAssertSucceeded();
                    AssertHandler.OnAssertCompleted -= assertHandler;
                }))
                .Verifiable();

            wrapperProvider.Setup(p => p.TypeLibWrapperFromProject(It.IsAny<string>()))
                            .Returns(typeLibMock.Object)
                            .Verifiable();

            typeLibMock.Setup(tlm => tlm.Dispose()).Verifiable();


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
                var engine = new TestEngine(state, fakesFactory.Object, interaction.Object, wrapperProvider.Object, dispatcher.Object, vbe);
                engine.TestCompleted += (source, args) => completionEvents.Add(args);
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun)
                {
                    Assert.Inconclusive("Parser Error");
                }
                engine.Run(engine.Tests);
            }
            Mock.Verify(dispatcher, interaction, wrapperProvider, typeLibMock);
            Assert.AreEqual(1, completionEvents.Count);
            Assert.AreEqual(new TestResult(TestOutcome.Succeeded), completionEvents.First().Result);
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Run_AndAssertInconclusive_RaisesCompletionEvent_Inconclusive()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var interaction = new Mock<IVBEInteraction>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var typeLibMock = new Mock<ITypeLibWrapper>();
            wrapperProvider.Setup(p => p.TypeLibWrapperFromProject(It.IsAny<string>()))
                            .Returns(typeLibMock.Object)
                            .Verifiable();
            long durationStub;
            interaction.Setup(ia => ia.RunTestMethod(typeLibMock.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out durationStub))
                .Callback(new RunTestMethodCallback((ITypeLibWrapper _, TestMethod method, EventHandler<AssertCompletedEventArgs> assertHandler, out long duration) =>
                {
                    duration = 0;
                    AssertHandler.OnAssertCompleted += assertHandler;
                    AssertHandler.OnAssertInconclusive("Test Message");
                    AssertHandler.OnAssertCompleted -= assertHandler;
                }))
                .Verifiable();

            typeLibMock.Setup(tlm => tlm.Dispose()).Verifiable();


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
                var engine = new TestEngine(state, fakesFactory.Object, interaction.Object, wrapperProvider.Object, dispatcher.Object, vbe);
                engine.TestCompleted += (source, args) => completionEvents.Add(args);
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun)
                {
                    Assert.Inconclusive("Parser Error");
                }
                engine.Run(engine.Tests);
            }
            Mock.Verify(dispatcher, interaction, wrapperProvider, typeLibMock);
            Assert.AreEqual(1, completionEvents.Count);
            Assert.AreEqual(new TestResult(TestOutcome.Inconclusive, "Test Message"), completionEvents.First().Result);
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Run_AndAssertFailed_RaisesCompletionEvent_Failed()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testMethods)
                .AddProjectToVbeBuilder();

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe);
            var interaction = new Mock<IVBEInteraction>();
            var wrapperProvider = new Mock<ITypeLibWrapperProvider>();
            var typeLibMock = new Mock<ITypeLibWrapper>();
            wrapperProvider.Setup(p => p.TypeLibWrapperFromProject(It.IsAny<string>()))
                            .Returns(typeLibMock.Object)
                            .Verifiable();

            long durationStub;
            interaction.Setup(ia => ia.RunTestMethod(typeLibMock.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out durationStub))
                .Callback(new RunTestMethodCallback((ITypeLibWrapper _, TestMethod method, EventHandler<AssertCompletedEventArgs> assertHandler, out long duration) =>
                {
                    duration = 0;
                    AssertHandler.OnAssertCompleted += assertHandler;
                    AssertHandler.OnAssertFailed("Test Message", "TestMethod1");
                    AssertHandler.OnAssertCompleted -= assertHandler;
                }))
                .Verifiable();

            typeLibMock.Setup(tlm => tlm.Dispose()).Verifiable();


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
                var engine = new TestEngine(state, fakesFactory.Object, interaction.Object, wrapperProvider.Object, dispatcher.Object, vbe);
                engine.TestCompleted += (source, args) => completionEvents.Add(args);
                parser.Parse(new CancellationTokenSource());
                if (!engine.CanRun)
                {
                    Assert.Inconclusive("Parser Error");
                }
                engine.Run(engine.Tests);
            }
            Mock.Verify(dispatcher, interaction, wrapperProvider, typeLibMock);
            Assert.AreEqual(1, completionEvents.Count);
            Assert.AreEqual(new TestResult(TestOutcome.Failed, string.Format(AssertMessages.Assert_FailedMessageFormat, "TestMethod1", "Test Message")), completionEvents.First().Result);
        }


        delegate void RunTestMethodCallback(ITypeLibWrapper wrapper, TestMethod method, EventHandler<AssertCompletedEventArgs> assertListener, out long duration);

        private const string TestModuleHeader = @"Option Explicit
Option Private Module

'@TestModule

Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Assert = CreateObject(""Rubberduck.AssertClass"")
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
    }
}
