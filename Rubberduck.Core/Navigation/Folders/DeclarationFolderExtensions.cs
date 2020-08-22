using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.JunkDrawer.Extensions;

namespace Rubberduck.Navigation.Folders
{
    public static class DeclarationFolderExtensions
    {
        public static string RootFolder(this Declaration declaration)
        {
            return declaration?.CustomFolder?.RootFolder() 
                   ?? declaration?.ProjectName 
                   ?? string.Empty;
        }

        public static bool IsInFolder(this Declaration declaration, string folder)
        {
            if (declaration?.CustomFolder is null || folder is null)
            {
                return false;
            }

            return declaration.CustomFolder.Equals(folder, StringComparison.Ordinal);
        }

        public static bool IsInSubFolder(this Declaration declaration, string folder)
        {
            var declarationFolder = declaration?.CustomFolder;
            if (declarationFolder is null || folder is null)
            {
                return false;
            }

            return declarationFolder.IsSubFolderOf(folder);
        }

        public static bool IsInFolderOrSubFolder(this Declaration declaration, string folder)
        {
            var declarationFolder = declaration?.CustomFolder;
            if (declarationFolder is null || folder is null)
            {
                return false;
            }

            return declaration.IsInFolder(folder)
                || declarationFolder.IsSubFolderOf(folder);
        }
    }
}
