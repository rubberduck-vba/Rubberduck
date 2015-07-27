namespace Rubberduck.Navigations.RegexSearchReplace
{
    public enum RegexSearchReplaceScope
    {
        Selection,
        CurrentBlock,
        CurrentFile,
        AllOpenedFiles,
        CurrentProject,
        EntireSolution
    }

    public interface IRegexSearchReplace
    {
        void Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void SearchAndReplace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
    }
}