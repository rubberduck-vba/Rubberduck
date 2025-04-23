using Rubberduck.UnitTesting;

namespace Rubberduck
{
    public class TestEngineProvider
    {
        private _Extension _extension;
        private ITestEngine _testEngine;
        public TestEngineProvider(_Extension extension)
        {
            _extension = extension;
        }
        public void SetTestEngine()
        {
            if (_testEngine != null)
            {
                return;
            }
            _testEngine = _extension.TestEngine;
        }
        public ITestEngine TestEngine { get => _testEngine; }
    }
}