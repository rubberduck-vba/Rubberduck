namespace Rubberduck.SourceControl
{
    public enum GitSettingsFile { Ignore, Attributes }

    public static class DefaultSettings
    {
        public static string GitAttributesText()
        {
            return "## gitattributes";
        }

        public static string GitIgnoreText()
        {
            return @"
###################
## Microsoft Office
###################

#UserForm Binary
*.frx

#Excel
*.xls
*.xlt
*.xlm
*.xla
*.xll
*.xlw
*.xlsx
*.xlsm
*.xltx
*.xltm
*.xlsb
*.xlam

#Access
*.ade
*.adp
*.adn
*.accdb
*.accde
*.accdr
*.accdt
*.accda
*.laccdb
*.mdb
*.cdb
*.mda
*.mdn
*.mdt
*.mdw
*.mdf
*.mde
*.mam
*.maq
*.mar
*.mat
*.maf
*.ldb

#Word
*.doc
*.dot
*.docx
*.docm
*.dotx
*.dotm
*.docb

#PowerPoint
*.ppt
*.pot
*.pps
*.pptx
*.pptm
*.potx
*.potm
*.ppam
*.ppsx
*.ppsm
*.sldx
*.sldm

#Publisher
*.pub
";
        }
    }
}
