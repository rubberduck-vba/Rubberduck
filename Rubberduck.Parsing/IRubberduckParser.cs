using NetOffice.VBIDEApi;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser : IParseResultProvider
    {
        /// <summary>
        /// Parses all code modules in specified project.
        /// </summary>
        /// <returns>Returns an <c>IParseTree</c> for each code module in the project; the qualified module name being the key.</returns>
        VBProjectParseResult Parse(VBProject vbProject, object owner = null);

        void RemoveProject(VBProject vbProject);

        void Parse(VBE vbe, object owner);
    }
}