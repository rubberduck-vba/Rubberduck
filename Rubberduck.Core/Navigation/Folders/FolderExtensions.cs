using System;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigation.Folders
{
    public static class FolderExtensions
    {
        public const char FolderDelimiter = '.';

        public static string RootFolder(this Declaration declaration)
        {
            return (declaration?.CustomFolder?.Trim('"') ?? string.Empty).Split(FolderDelimiter).FirstOrDefault() 
                   ?? declaration?.ProjectName 
                   ?? string.Empty;
        }

        public static string SubFolderRoot(this string folder, string subfolder)
        {
            if (folder is null || subfolder is null || !folder.Trim('"').StartsWith(subfolder.Trim('"')))
            {
                return string.Empty;
            }

            var subPath = folder.Trim('"').Substring(subfolder.Length + 1);
            return subPath.Split(FolderDelimiter).FirstOrDefault() ?? string.Empty;
        }

        public static bool IsInFolder(this Declaration declaration, string folder)
        {
            if (declaration?.CustomFolder is null || folder is null)
            {
                return false;
            }

            return declaration.CustomFolder.Trim('"').Equals(folder.Trim('"'), StringComparison.Ordinal);
        }

        public static bool IsInSubFolder(this Declaration declaration, string folder)
        {
            if (declaration?.CustomFolder is null || folder is null)
            {
                return false;
            }

            var folderPath = folder.Trim('"').Split(FolderDelimiter);
            var declarationPath = declaration.CustomFolder.Trim('"').Split(FolderDelimiter);

            if (folderPath.Length >= declarationPath.Length)
            {
                return false;
            }

            return declarationPath.Take(folderPath.Length).SequenceEqual(folderPath, StringComparer.Ordinal);
        }

        public static bool IsInFolderOrSubFolder(this Declaration declaration, string folder)
        {
            if (declaration?.CustomFolder is null || folder is null)
            {
                return false;
            }

            var folderPath = folder.Trim('"').Split(FolderDelimiter);
            var declarationPath = declaration.CustomFolder.Trim('"').Split(FolderDelimiter);

            for (var depth = 0; depth < folderPath.Length && depth < declarationPath.Length; depth++)
            {
                if (!folderPath[depth].Equals(declarationPath[depth], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
