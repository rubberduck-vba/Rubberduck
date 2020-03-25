using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    internal class MockedTestEngine : IDisposable
    {
        public delegate void RunTestMethodCallback(ITypeLibWrapper wrapper, TestMethod method, EventHandler<AssertCompletedEventArgs> assertListener, out long duration);
        public delegate (long Duration, TestResult result) ReturnTestResult(ITypeLibWrapper wrapper, TestMethod method, EventHandler<AssertCompletedEventArgs> assertListener, out long duration);

        private const string TestMethodTemplate = 
@"'@TestMethod
Public Sub TestMethod{0}()
End Sub";

        private const string IgnoredTestTemplate =
@"'@TestMethod
'@IgnoreTest
Public Sub TestMethod{0}()
End Sub";

        private const string TestMethodCategoryTemplate =
            @"'@TestMethod(""{1}"")
Public Sub TestMethod{0}()
End Sub";

        private const string IgnoredTestCategoryTemplate =
            @"'@TestMethod(""{1}"")
'@IgnoreTest
Public Sub TestMethod{0}()
End Sub";

        private readonly Mock<IFakesFactory> _fakesFactory = new Mock<IFakesFactory>();
        private readonly Mock<IFakes> _createdFakes = new Mock<IFakes>();
        private long _durationStub;

        private MockedTestEngine()
        {
            Dispatcher.Setup(d => d.InvokeAsync(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke())
                .Verifiable();
            Dispatcher.Setup(d => d.StartTask(It.IsAny<Action>(), It.IsAny<TaskCreationOptions>()))
                .Returns((Action action, TaskCreationOptions options) =>
                    {
                        action.Invoke();
                        return Task.CompletedTask;
                    })
                .Verifiable();

            TypeLib.Setup(tlm => tlm.Dispose()).Verifiable();
            WrapperProvider.Setup(p => p.TypeLibWrapperFromProject(It.IsAny<string>())).Returns(TypeLib.Object).Verifiable();

            _fakesFactory.Setup(factory => factory.Create()).Returns(_createdFakes.Object);
        }

        public MockedTestEngine(string testModuleCode) : this()
        {
            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, TestModuleHeader + testModuleCode)
                .AddProjectToVbeBuilder();

            Vbe = builder.Build();
            ParserState = MockParser.Create(Vbe.Object).State;
            TestEngine = new SynchronouslySuspendingTestEngine(ParserState, _fakesFactory.Object, VbeInteraction.Object, WrapperProvider.Object, Dispatcher.Object, Vbe.Object, ParserState.ProjectsProvider);
        }

        public MockedTestEngine(IReadOnlyList<string> moduleNames, IReadOnlyList<int> methodCounts) : this()
        {
            if (moduleNames.Count != methodCounts.Count)
            {
                Assert.Inconclusive("Test setup error.");
            }

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);

            for (var index = 0; index < moduleNames.Count; index++)
            {
                var testModuleCode = string.Join(Environment.NewLine, Enumerable.Range(1, methodCounts[index]).Select(num => GetTestMethod(num)));

                project.AddComponent(moduleNames[index], ComponentType.StandardModule, TestModuleHeader + testModuleCode);
            }

            project.AddProjectToVbeBuilder();
            Vbe = builder.Build();
            ParserState = MockParser.Create(Vbe.Object).State;
            TestEngine = new SynchronouslySuspendingTestEngine(ParserState, _fakesFactory.Object, VbeInteraction.Object, WrapperProvider.Object, Dispatcher.Object, Vbe.Object, ParserState.ProjectsProvider);
        }

        public MockedTestEngine(int testMethodCount) 
            : this(string.Join(Environment.NewLine, Enumerable.Range(1, testMethodCount).Select(num => GetTestMethod(num))))
        { }

        public MockedTestEngine(List<(TestOutcome Outcome, string Output, long duration)> results) 
            : this(string.Join(Environment.NewLine, Enumerable.Range(1, results.Count).Select(num => GetTestMethod(num, results[num - 1].Outcome == TestOutcome.Ignored))))
        {
            ParserState.OnParseRequested(this);
            var testMethodCount = results.Count;
            var testMethods = TestEngine.Tests.ToList();

            if (testMethods.Count != testMethodCount)
            {
                Assert.Inconclusive("Test setup failure.");
            }

            for (var test = 0; test < results.Count; test++)
            {
                var (outcome, output, duration) = results[test];
                SetupAssertCompleted(testMethods[test], new TestResult(outcome, output, duration));
            }
        }

        public RubberduckParserState ParserState { get; set; }

        public Mock<IVBE> Vbe { get; set; }

        public ITestEngine TestEngine { get; set; }

        public Mock<IVBEInteraction> VbeInteraction { get; } = new Mock<IVBEInteraction>();

        public Mock<ITypeLibWrapper> TypeLib { get; } = new Mock<ITypeLibWrapper>();

        public Mock<IUiDispatcher> Dispatcher { get; } = new Mock<IUiDispatcher>();

        public Mock<ITypeLibWrapperProvider> WrapperProvider { get; } = new Mock<ITypeLibWrapperProvider>();

        public static string GetTestMethod(int number, bool ignored = false, string category = null) =>
            category is null 
                ? string.Format(ignored ? IgnoredTestTemplate : TestMethodTemplate, number) 
                : string.Format(ignored ? IgnoredTestCategoryTemplate : TestMethodCategoryTemplate, number, category);

        public void SetupAssertCompleted(Action action)
        {
            if (action is null)
            {
                VbeInteraction.Setup(ia => ia.RunTestMethod(TypeLib.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out _durationStub))
                    .Verifiable();
                return;
            }

            var callback = GetAssertCompletedCallback(action);

            VbeInteraction
                .Setup(ia => ia.RunTestMethod(TypeLib.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out _durationStub))
                .Callback(callback)
                .Verifiable();
        }

        [SuppressMessage("ReSharper", "ImplicitlyCapturedClosure")] //This is fine - the closures' lifespan is limited to the test.
        public void SetupAssertCompleted(TestMethod testMethod, TestResult result)
        {
            Action action;
            switch (result.Outcome)
            {
                // Tests involving this case have concurrency issues when the UIDispatcher is mocked. The callbacks in production
                // are moderated by flushing the message queue. There is no way to mock this behavior because in the testing environment
                // there is no STA boundary being crossed, so the internal implementation of the AssertHandler is not constrained by
                // the rental context that would normally be managed by Interop. These tests *do* pass as of the commit that adds them,
                // but only if the context is force by running them through the debugger. Leaving these in as commented code mainly for
                // the purpose of documenting this. Single asserts are fine. Update - added spin wait instead. This may still be a FIXME?

                case TestOutcome.Failed:
                    action = () =>
                    {
                        var assert = new AssertClass();
                        assert.Fail(result.Output);
                    };
                    break;
                case TestOutcome.Inconclusive:
                    action = () =>
                    {
                        var assert = new AssertClass();
                        assert.Inconclusive(result.Output);
                    };
                    break;
                case TestOutcome.Succeeded:
                    action = () =>
                    {
                        var assert = new AssertClass();
                        assert.Succeed();
                    };
                    break;
                default:
                    action = () => { };
                    break;
            }

            var callback = new RunTestMethodCallback((ITypeLibWrapper _, TestMethod method, EventHandler<AssertCompletedEventArgs> assertHandler, out long duration) =>
            {
                duration = 0;
                AssertHandler.OnAssertCompleted += assertHandler;
                action.Invoke();
                AssertHandler.OnAssertCompleted -= assertHandler;
            });

            VbeInteraction
                .Setup(ia => ia.RunTestMethod(TypeLib.Object, It.Is<TestMethod>(test => testMethod.Equals(test)), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out _durationStub))
                .Callback(callback)
                .Verifiable();
        }

        public RunTestMethodCallback GetAssertCompletedCallback(Action action)
        {
            return (ITypeLibWrapper _, TestMethod method, EventHandler<AssertCompletedEventArgs> assertHandler, out long duration) =>
            {
                duration = 0;
                AssertHandler.OnAssertCompleted += assertHandler;
                action.Invoke();
                AssertHandler.OnAssertCompleted -= assertHandler;
            };
        }

        public void Dispose()
        {
            ParserState?.Dispose();
        }

        public const string TestModuleHeader = @"Option Explicit
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

        private class SynchronouslySuspendingTestEngine : TestEngine
        {
            private readonly RubberduckParserState _state;

            public SynchronouslySuspendingTestEngine(
                RubberduckParserState state,
                IFakesFactory fakesFactory,
                IVBEInteraction declarationRunner,
                ITypeLibWrapperProvider wrapperProvider,
                IUiDispatcher uiDispatcher,
                IVBE vbe,
                IProjectsProvider projectsProvider)
                : base(state, fakesFactory, declarationRunner, wrapperProvider, uiDispatcher, vbe, projectsProvider)
            {
                _state = state;
            }

            protected override void RunInternal(IEnumerable<TestMethod> tests)
            {
                if (!CanRun)
                {
                    return;
                }
                //We have to do this on the same thread here to guarantee that the actions runs before the assert in the unit tests is called.
                _state.OnSuspendParser(this, AllowedRunStates, () => RunWhileSuspended(tests));
            }
        }
    }
}
