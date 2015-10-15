using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI.ParserProgress
{
    public class ParsingProgressPresenter
    {
        private readonly IRubberduckParser _parser;

        public ParsingProgressPresenter(IRubberduckParser parser)
        {
            _parser = parser;
        }

        public VBProjectParseResult Parse(VBProject project)
        {
            var viewModel = new ParserProgessViewModel(_parser, project);
            using (var view = new ProgressDialog(viewModel))
            {
                view.ShowDialog();
                return view.Result;
            }
        }
    }
}
