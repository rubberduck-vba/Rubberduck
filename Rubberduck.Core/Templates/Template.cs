using System.Text.RegularExpressions;

namespace Rubberduck.Templates
{
    /// <remarks>
    /// Template can be either built-in or user-defined. For a built-in template, the
    /// metadata should be stored in the <see cref="Resources.Templates"/>
    /// resource, with specific entries, currently Name, Caption, Description and Code.
    /// Due to the fact that we cannot strong-type the reference to the resource entries
    /// the class has unit tests to validate that the crucial elements are present in the
    /// resource to guard against runtime errors/unexpected behavior due to missing/malformed
    /// entries in the resources. 
    /// </remarks>
    public class Template : ITemplate
    {
        public const string TemplateExtension = ".template";
        public static int ExtensionLength = TemplateExtension.Length;

        private readonly ITemplateFileHandler _handler;
        private readonly string _builtInCode;

        public Template(string name, ITemplateFileHandler handler)
        {
            _handler = handler;
            
            Name = name.EndsWith(TemplateExtension) ? name.Substring(0, name.Length - ExtensionLength) : name;
            _builtInCode = Resources.Templates.ResourceManager.GetString(Name + "_Code");
            IsUserDefined = string.IsNullOrWhiteSpace(_builtInCode);

            if (IsUserDefined)
            {
                var code = handler.Read();
                if (!string.IsNullOrWhiteSpace(code))
                {
                    var regex = new Regex(@"^Attribute VB_Ext_KEY\s+=\s+""Rubberduck(Caption|Description)"",\s+""(.+)""", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.CultureInvariant);
                    var matches = regex.Matches(code);
                    foreach (Match match in matches)
                    {
                        switch (match.Groups[1].Value)
                        {
                            case "Caption":
                                Caption = match.Groups[2].Value;
                                break;
                            case "Description":
                                Description = match.Groups[2].Value;
                                break;
                        }
                    }
                }

                if (string.IsNullOrEmpty(Caption)) Caption = Name;
                if(string.IsNullOrEmpty(Description)) Description = Name;
            }
            else
            {
                VerifyFile(Name, handler);
                (Caption, Description) = GetBuiltInMetaData(Name);
            }
        }

        public string Name { get; }
        public bool IsUserDefined { get; }
        public string Caption { get; }
        public string Description { get; }

        public string Read() => _handler.Read() ?? _builtInCode;
        
        public void Write(string content) => _handler.Write(content);

        private static void VerifyFile(string name, ITemplateFileHandler handler)
        {
            if (handler.Exists)
            {
                return;
            }

            var content = Resources.Templates.ResourceManager.GetString(name + "_Code");
            handler.Write(content);
        }

        private static (string caption, string description) GetBuiltInMetaData(string name)
        {
            var caption = Resources.Templates.ResourceManager.GetString(name + "_Caption");
            var description = Resources.Templates.ResourceManager.GetString(name + "_Description");

            return (caption, description);
        }
    }
}