namespace RubberduckTests.UnitTesting
{
    //[TestClass]
    //public class EngineTests
    //{
    //    private TestEngine _engine;
    //    private Mock<IHostApplication> _hostAppMock;
    //    private readonly QualifiedModuleName _moduleName = new QualifiedModuleName("VBAProject", "TestModule1");

    //    private TestMethod _successfulMethod;
    //    private TestMethod _failedMethod;
    //    private TestMethod _inconclusiveMethod;
    //    private TestMethod _notRunMethod;

    //    private bool _wasEventRaised;
    //    private int _eventCount;

    //    [TestInitialize]
    //    public void Initialize()
    //    {
    //        _wasEventRaised = false;
    //        _eventCount = 0;

    //        _engine = new TestEngine();
    //        _hostAppMock = new Mock<IHostApplication>();

    //        _successfulMethod = new TestMethod(new QualifiedMemberName(_moduleName, "TestMethod1"), _hostAppMock.Object);
    //        _failedMethod = new TestMethod(new QualifiedMemberName(_moduleName, "TestMethod2"), _hostAppMock.Object);
    //        _inconclusiveMethod = new TestMethod(new QualifiedMemberName(_moduleName, "TestMethod3"), _hostAppMock.Object);
    //        _notRunMethod = new TestMethod(new QualifiedMemberName(_moduleName, "TestMethod4"), _hostAppMock.Object);

    //        var tests = new Dictionary<TestMethod, TestResult>
    //        {
    //            {_successfulMethod, new TestResult(TestOutcome.Succeeded)},
    //            {_failedMethod, new TestResult(TestOutcome.Failed)},
    //            {_inconclusiveMethod, new TestResult(TestOutcome.Inconclusive)},
    //            {_notRunMethod, null}
    //        };

    //        _engine.AllTests = tests;
    //    }

    //    [TestMethod]
    //    public void TestEngine_FailedTests()
    //    {
    //        var actual = _engine.FailedTests().First();

    //        Assert.AreEqual(_failedMethod, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_SuccessfulTests()
    //    {
    //        var actual = _engine.PassedTests().First();

    //        Assert.AreEqual(_successfulMethod, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_NotRunTests()
    //    {
    //        var actual = _engine.NotRunTests().First();

    //        Assert.AreEqual(_notRunMethod, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_LastRunTests_ReturnsAllRunTests()
    //    {
    //        var actual = _engine.LastRunTests().ToList();
    //        var expected = new List<TestMethod>()
    //        {
    //            _failedMethod, _inconclusiveMethod, _successfulMethod
    //        };

    //        CollectionAssert.AreEquivalent(expected, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_LastRunTests_Successful()
    //    {
    //        var actual = _engine.LastRunTests(TestOutcome.Succeeded).First();

    //        Assert.AreEqual(_successfulMethod, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_LastRunTests_Failed()
    //    {
    //        var actual = _engine.LastRunTests(TestOutcome.Failed).First();

    //        Assert.AreEqual(_failedMethod, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_LastRunTests_Inconclusive()
    //    {
    //        var actual = _engine.LastRunTests(TestOutcome.Inconclusive).First();

    //        Assert.AreEqual(_inconclusiveMethod, actual);
    //    }

    //    [TestMethod]
    //    public void TestEngine_Run_ModuleIntialize_IsRunOnce()
    //    {
    //        //arrange
    //        _engine.ModuleInitialize += CatchEvent;

    //        var tests = _engine.AllTests.Keys;

    //        //act
    //        _engine.Run(tests);

    //        Assert.IsTrue(_wasEventRaised, "Module Intialize was not run.");
    //        Assert.AreEqual(1, _eventCount, "Module Intialzie expected to be run once.");
    //    }

    //    [TestMethod]
    //    public void TestEngine_Run_ModuleCleanup_IsRunOnce()
    //    {
    //        //arrange
    //        _engine.ModuleCleanup += CatchEvent;

    //        //act
    //        _engine.Run(_engine.AllTests.Keys);

    //        //assert
    //        Assert.IsTrue(_wasEventRaised, "Module Cleanup was not run.");
    //        Assert.AreEqual(1, _eventCount, "Module Cleanup expected to be run once.");
    //    }

    //    [TestMethod]
    //    public void TestEngine_Run_MethodIntialize_IsRunForEachTestMethod()
    //    {
    //        //arrange
    //        var expectedCount = _engine.AllTests.Count;
    //        _engine.MethodInitialize += CatchEvent;

    //        //act
    //        _engine.Run(_engine.AllTests.Keys);

    //        //assert
    //        Assert.IsTrue(_wasEventRaised, "Method Intialize was not run.");
    //        Assert.AreEqual(expectedCount, _eventCount, "Method Intialized was expected to be run {0} times", expectedCount);
    //    }

    //    [TestMethod]
    //    public void TestEngine_Run_MethodCleanup_IsRunForEachTestMethod()
    //    {
    //        //arrange
    //        var expectedCount = _engine.AllTests.Count;
    //        _engine.MethodCleanup += CatchEvent;

    //        //act
    //        _engine.Run(_engine.AllTests.Keys);

    //        //assert
    //        Assert.IsTrue(_wasEventRaised, "Method Initialize was not run.");
    //        Assert.AreEqual(expectedCount, _eventCount, "Method Initialized was expected to be run {0} times", expectedCount);
    //    }

    //    [TestMethod]
    //    public void TestEngine_Run_TestCompleteIsRaisedForEachTestMethod()
    //    {
    //        //arrange
    //        var expectedCount = _engine.AllTests.Count;
    //        _engine.TestCompleted += EngineOnTestComplete;

    //        //act
    //        _engine.Run(_engine.AllTests.Keys);

    //        //assert
    //        Assert.IsTrue(_wasEventRaised, "TestCompleted event was not raised.");
    //        Assert.AreEqual(expectedCount, _eventCount, "TestCompleted event was expected to be raised {0} times.", expectedCount);
    //    }

    //    [TestMethod]
    //    public void TestEngine_Run_WhenTestListIsEmpty_Bail()
    //    {
    //        //arrange 
    //        _engine.MethodInitialize += CatchEvent;

    //        //act
    //        _engine.Run(new List<TestMethod>());

    //        //assert
    //        Assert.IsFalse(_wasEventRaised, "No methods should run when passed an empty list of tests.");
    //    }

    //    private void EngineOnTestComplete(object sender, TestCompletedEventArgs testCompletedEventArgs)
    //    {
    //        CatchEvent();
    //    }

    //    private void CatchEvent(object sender, TestModuleEventArgs e)
    //    {
    //        CatchEvent();
    //    }

    //    private void CatchEvent()
    //    {
    //        _wasEventRaised = true;
    //        _eventCount++;
    //    }
    //}
}
