using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public static class CommandBarControlCaptionGuard
    {
        private static int MAX_CAPTION_LENGTH = 255;
        private static int MODIFIED_PRE_POST_ELLIPSIS_LENGTH = 30;
        private static string ELLIPSIS = "...";


        public static string ApplyGuard(string proposedCaption)
        {
            if (proposedCaption == null || proposedCaption.Length < MAX_CAPTION_LENGTH)
            {
                return proposedCaption;
            }

            if (IsMethodFormat(proposedCaption))
            {
                var splitCaption = proposedCaption.Split(new[] { "." }, System.StringSplitOptions.None);
                Debug.Assert(splitCaption.Length == 3);
                //splitCaption[0] = Coordinates plus filename (allowed to be nearly 200 characters)
                //splitCaption[1] = Module Name (limited to 31 characters by VBE)
                //splitCaption[2] = Sub or Function name plus a type string at the end (e.g. "(procedure)")

                //Reduce the filename first if it is too 'long'
                splitCaption[0] = MinimizeCaptionPortionName(splitCaption[0]);

                if ( splitCaption[0].Length + splitCaption[1].Length + splitCaption[2].Length > MAX_CAPTION_LENGTH)
                {
                    //still too long, truncate the method name
                    splitCaption[2] = MinimizeCaptionPortionName(splitCaption[2]);
                }

                return $"{splitCaption[0]}.{splitCaption[1]}.{splitCaption[2]}";
            }

            //Don't recognize the format, so bluntly truncate and avoid the exception
            return proposedCaption.Substring(0, MAX_CAPTION_LENGTH - (ELLIPSIS.Length * 2)) + ELLIPSIS;
        }

        public static bool IsMethodFormat(string proposedCaption)
        {
            //Example of a caption input when a method is selected: 
            //"L23C13 | TheFilename.TheModuleName.TheMethodName (procedure)"
            const string pattern = @"^[L][0-9]+[C][0-9]+\s[|]\s[a-zA-Z0-9]+[.][a-zA-Z0-9]+[.][a-zA-Z0-9]+\s[(][a-z]+[)]\z";
            return Regex.IsMatch(proposedCaption, pattern);
        }

        private static string MinimizeCaptionPortionName(string methodNamePlusTypeIdentifier)
        {
            if (methodNamePlusTypeIdentifier.Length > MODIFIED_PRE_POST_ELLIPSIS_LENGTH * 2 + ELLIPSIS.Length)
            {
                var preEllipsis = methodNamePlusTypeIdentifier.Substring(0, MODIFIED_PRE_POST_ELLIPSIS_LENGTH);
                var postEllipsis = methodNamePlusTypeIdentifier.Substring(methodNamePlusTypeIdentifier.Length - MODIFIED_PRE_POST_ELLIPSIS_LENGTH);
                return $"{preEllipsis}{ELLIPSIS}{postEllipsis}";
            }
            return methodNamePlusTypeIdentifier;
        }
    }
}
