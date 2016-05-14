using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// A TestExplorer model that discovers unit tests in standard modules (.bas) marked with a '@TestModule marker.
    /// </summary>
    public class StandardModuleTestExplorerModel : TestExplorerModelBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;

        public StandardModuleTestExplorerModel(VBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;
        }

        public override void Refresh()
        {
            IsBusy = true;

            var tests = UnitTestHelpers.GetAllTests(_vbe, _state);

            ClearLastRun();
            ExecutedCount = 0;
            foreach (var test in tests)
            {                
                AddExecutedTest(test);
            }

            // ReSharper disable once ExplicitCallerInfoArgument
            OnPropertyChanged("Tests");
            IsBusy = false;
        }
    }
}