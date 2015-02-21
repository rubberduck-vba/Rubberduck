using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection
{
    public enum MemberType
    {
        Field,
        Sub,
        Function,
        PropertyLet,
        PropertyGet,
        PropertySet
    }

    public enum MemberVisibility
    {
        Public,
        Private,
        Friend
    }

    public class Member
    {
        public Member(MemberVisibility visibility, 
                      MemberType memberType, 
                      string projectName,
                      string moduleName,
                      string name, 
                      string signature, 
                      string[] body, 
                      IEnumerable<MemberAttribute> attributes)
        {
            _projectName = projectName;
            _moduleName = moduleName;
            _memberType = memberType;
            _name = name;
            _signature = signature;
            _body = body;
            _attributes = attributes.ToDictionary((attribute) => attribute.Name, (attribute) => attribute);
        }

        private readonly string _projectName;
        private readonly string _moduleName;
        public string QualifiedName { get { return string.Concat(_projectName, ".", _moduleName, ".", _name); } }

        private readonly MemberType _memberType;
        public MemberType MemberType { get { return _memberType; } }

        private readonly MemberVisibility _memberVisibility;
        public MemberVisibility MemberVisibility { get { return _memberVisibility; } }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _signature;
        public string Signature { get { return _signature; } }

        private readonly string[] _body;
        public string[] Body { get { return _body; } }

        private readonly IDictionary<string, MemberAttribute> _attributes;
        public bool HasAttribute(string name)
        {
            return _attributes.ContainsKey(name);
        }

        public bool HasAttribute<TAttribute>() where TAttribute : MemberAttributeBase, new()
        {
            return HasAttribute(new TAttribute().Name);
        }

        public MemberAttribute GetAttribute(string name)
        {
            return _attributes[name];
        }

        public MemberAttribute GetAttribute<TAttribute>() where TAttribute : MemberAttributeBase, new()
        {
            return GetAttribute(new TAttribute().Name);
        }

        private static readonly string[] _keywords = 
            new[] { "Dim ", "Private ", "Public ", "Friend ", 
                    "Function ", "Sub ", "Property ", 
                    "Static ", "Const ", "Global ", "Enum ", "Type ", "Declare " };

        public static bool TryParse(string[] body, string projectName, string moduleName, out Member result)
        {
            var signature = body.FirstOrDefault(line => _keywords.Any(keyword => line.StartsWith(keyword)));
            if (signature == null)
            {
                result = null;
                return false;
            }

            signature = signature.Replace("\r", string.Empty);
            body = body.Select(line => line.Replace("\r", string.Empty)).ToArray();

            var withoutKeyword = signature.Substring((_keywords.First(keyword => signature.StartsWith(keyword))).Length);
            var name = withoutKeyword.Split(' ')[1]
                                     .Split('(')[0];

            var type = GetMemberType(signature, withoutKeyword);
            var visibility = GetMemberVisibility(signature);

            var signatureLineIndex = body.ToList().IndexOf(signature);
            var attributes = MemberAttribute.GetAttributes(body.Take(signatureLineIndex)
                                                               .Where(line => line.StartsWith("'@")));

            result = new Member(visibility, type, projectName, moduleName, name, signature, body, attributes);
            return true;
        }

        private static MemberVisibility GetMemberVisibility(string signature)
        {
            MemberVisibility result;
            if (signature.StartsWith("Dim ") || 
                signature.StartsWith("Private ") || 
                signature.StartsWith("Static ") ||
                signature.StartsWith("Const "))
            {
                result = MemberVisibility.Private;
            }
            else if (signature.StartsWith("Friend"))
            {
                result = MemberVisibility.Friend;
            }
            else
            {
                result = MemberVisibility.Public;
            }
            return result;
        }

        private static MemberType GetMemberType(string signature, string withoutKeyword)
        {
            MemberType result;
            if (withoutKeyword.StartsWith("Function ") || signature.StartsWith("Function "))
            {
                result = MemberType.Function;
            }
            else if (withoutKeyword.StartsWith("Sub ") || signature.StartsWith("Sub "))
            {
                result = MemberType.Sub;
            }
            else if (withoutKeyword.StartsWith("Property Get") || signature.StartsWith("Property Get "))
            {
                result = MemberType.PropertyGet;
            }
            else if (withoutKeyword.StartsWith("Property Let") || signature.StartsWith("Property Let "))
            {
                result = MemberType.PropertyLet;
            }
            else if (withoutKeyword.StartsWith("Property Set") || signature.StartsWith("Property Set "))
            {
                result = MemberType.PropertySet;
            }
            else
            {
                result = MemberType.Field;
            }
            return result;
        }
    }

    public class MemberAttribute
    {
        private const string AttributeSyntax = 
            @"\'\@(?<AttributeName>\w+)(?<Parameters>\(((?<Parameter>(\""[a-zA-Z][a-zA-Z_0-9]*\"")|([0-9]+(\.[0-9]+)?))(\,\s*)?)+\))?";

        public MemberAttribute(string name, IEnumerable<string> parameters)
        {
            _name = name;
            _parameters = parameters;
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly IEnumerable<string> _parameters;
        public IEnumerable<string> Parameters { get { return _parameters; } }

        public static bool TryParse(string codeLine, out MemberAttribute result)
        {
            var match = Regex.Match(codeLine, AttributeSyntax);

            if (match.Success)
            {
                var name = match.Groups["AttributeName"].Value;

                IEnumerable<string> parameters = null;
                if (match.Groups["Parameters"].Success)
                {
                    parameters = match.Groups["Parameter"].Captures
                                                          .Cast<Group>()
                                                          .Select(group => group.Value)
                                                          .Select(value => value.StartsWith("\"") && value.EndsWith("\"") 
                                                                                    ? value.Substring(1, value.Length - 2) 
                                                                                    : value)
                                                          .ToList();
                }

                result = new MemberAttribute(name, parameters);
                return true;
            }
            else
            {
                result = null;
                return false;
            }
        }

        public static IEnumerable<MemberAttribute> GetAttributes(IEnumerable<string> codeLines)
        {
            foreach (var line in codeLines)
            {
                MemberAttribute attribute;
                if (MemberAttribute.TryParse(line, out attribute))
                {
                    yield return attribute;
                }
            }
        }
    }
}
