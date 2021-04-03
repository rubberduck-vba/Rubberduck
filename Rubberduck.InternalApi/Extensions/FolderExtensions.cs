using System.Linq;

namespace Rubberduck.InternalApi.Extensions
{
    public static class FolderExtensions
    {
        public const char FolderDelimiter = '.';

        public static string RootFolder(this string folder)
        {
            return (folder ?? string.Empty).Split(FolderExtensions.FolderDelimiter).FirstOrDefault();
        }

        public static string SubFolderName(this string folder)
        {
            return (folder ?? string.Empty).Split(FolderExtensions.FolderDelimiter).LastOrDefault();
        }

        public static string SubFolderPathRelativeTo(this string subFolder, string folder)
        {
            if (subFolder is null || folder is null)
            {
                return string.Empty;
            }

            if (folder.Length == 0)
            {
                return subFolder;
            }

            if (!subFolder.StartsWith(folder))
            {
                return string.Empty;
            }

            return subFolder.Substring(folder.Length + 1);
        }

        public static string SubFolderRoot(this string subFolder, string folder)
        {
            var subPath = subFolder?.SubFolderPathRelativeTo(folder) ?? string.Empty;
            return subPath.Split(FolderDelimiter).FirstOrDefault() ?? string.Empty;
        }

        public static string ParentFolder(this string folder)
        {
            if (folder is null || !folder.Contains(FolderDelimiter))
            {
                return string.Empty;
            }

            var lastDelimiterIndex = folder.LastIndexOf(FolderDelimiter);
            return folder.Substring(0, lastDelimiterIndex);
        }

        public static bool IsSubFolderOf(this string subFolder, string folder)
        {
            return subFolder != null
                   && folder != null
                   && folder.Length < subFolder.Length
                   && subFolder.StartsWith(folder)
                   && subFolder[folder.Length] == FolderDelimiter;
        }
    }
}