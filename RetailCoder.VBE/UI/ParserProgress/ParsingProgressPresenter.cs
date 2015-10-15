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
            using (var view = new ProgressDialog(_parser, project))
            {
                view.ShowDialog();
                return view.Result;
            }
        }
    }
}
