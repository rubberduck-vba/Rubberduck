namespace Rubberduck.Navigations.RegexSearchReplace
{
    public interface IRegexSearchReplaceView
    {
        string SearchPattern { get; }
        string ReplacePattern { get; }
        RegexSearchReplaceScope Scope { get; }
    }
}