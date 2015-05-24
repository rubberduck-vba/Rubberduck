using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI
{
    public class ParsingProgressPresenter
    {
        public VBProjectParseResult Parse(IRubberduckParser parser, VBProject project)
        {
            using (var view = new ProgressDialog(parser, project))
            {
                view.ShowDialog();
                return view.Result;
            }
        }
    }
}
