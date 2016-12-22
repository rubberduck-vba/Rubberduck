using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Extensions
{
    public static class MSAccessComponentTypeExtensions
    {
        internal const string AccessFormExtension = ".accfrm";
        internal const string AccessReportExtension = ".accrpt";

        /// <summary>
        /// Returns the proper file extension for the MS Access Component Type.
        /// </summary>
        /// <param name="componentType"></param>
        /// <returns>File extension that includes a preceeding "dot" (.) </returns>
        public static string FileExtension(this MSAccessComponentType componentType)
        {
            switch (componentType)
            {
                case MSAccessComponentType.Form:
                    return AccessFormExtension;
                case MSAccessComponentType.Report:
                    return AccessReportExtension;
                default:
                    return string.Empty;
            }
        }
    }
}
