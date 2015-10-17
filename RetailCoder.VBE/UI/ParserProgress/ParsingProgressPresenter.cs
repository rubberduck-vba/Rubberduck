using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI.ParserProgress
{
    public interface IParsingProgressPresenter
    {
        VBProjectParseResult Parse(VBProject project);
    }

    public class ParsingProgressPresenter : IParsingProgressPresenter
    {
        private readonly IRubberduckParser _parser;

        public ParsingProgressPresenter(IRubberduckParser parser)
        {
            _parser = parser;
        }

        public VBProjectParseResult Parse(VBProject project)
        {
            using (var view = new ProgressDialog(new ParserProgessViewModel(_parser, project)))
            {
                view.ShowDialog();
                return view.Result;
            }
        }
    }
}
