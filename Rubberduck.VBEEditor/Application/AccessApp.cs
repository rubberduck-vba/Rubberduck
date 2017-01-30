using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Access;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Application
{
    public class AccessApp : HostApplicationBase<Microsoft.Office.Interop.Access.Application>
    {
        public AccessApp() : base("Access") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call);
        }

        public List<string> FormDeclarations(QualifiedModuleName qualifiedModuleName)
        {
            //TODO: Determine if component is Form/Report
            string filePath = Path.Combine(ExportPath, qualifiedModuleName.Name +  MSAccessComponentType.Form.FileExtension());
            Application.SaveAsText(AcObjectType.acForm, qualifiedModuleName.Name, filePath);
            var code = File.ReadAllText(filePath);
            File.Delete(filePath);


            bool isMainDocumentSection = false;
            bool isInControlsSection = false;
            bool isTargetProperty = false;

            string currentBlockName = "";
            string currPropName = "";
            string currPropValue = "";
            string valuePart = "";

            string[] propertyNameTargets = new string[] {"Name", "Class", "OnLoad", "OLEClass", "EventProcPrefix", "ControlSource"};
            string[] documentSectionNames = new string[] {"FormHeader", "PageHeader", "BreakHeader", "Section", "BreakFooter", "PageFooter", "FormFooter"}
            Dictionary<string,string> documentSectionNames = new Dictionary<string,string>();


            const string documentTypeName = "Form";

            const string PROPERTY_NAME_VALUE_DELIMITER = " =";
            const string HEX_START_OF_LINE = "0x";
            const string HEX_END_OF_LINE = " ,";


            string[] lines = code.Split(System.Environment.NewLine.ToCharArray());
            foreach (string rawLine in lines)
            {
                if (rawLine.Length > 0) //skip blank lines
                {
                    string line = rawLine.Trim();
                    int indent = rawLine.Length - line.Length;

                    if (line.StartsWith("Begin")) //Start of a Control Group - Nothing to do here ***
                    {
                    }

                    else if (line.StartsWith("Begin ")) //Start of a Section or Control ***
                    {
                        currentBlockName = line.Substring("Begin ".Length + 1);

                        isMainDocumentSection = currentBlockName == documentTypeName;
                        if (!isInControlsSection)
                        {
                            isInControlsSection = documentSectionNames.ContainsKey(currentBlockName);
                        }

                    }

                    else if (line == "End") //End of a section, control or control group ***
                    {
                        if (indent == 0)
                        {
                            //We've reached the end of the document text, what follows is the VBA (which is parsed elsewhere)
                            break;
                        }
                    }

                    else if (line.StartsWith(HEX_START_OF_LINE) && isTargetProperty) //Hex property ***
                    {
                        if (line.EndsWith(HEX_END_OF_LINE)) 
                        {
                            valuePart = line.Substring(HEX_START_OF_LINE.Length, line.Length - HEX_START_OF_LINE.Length - HEX_END_OF_LINE.Length);
                        } 
                        else{
                            valuePart = line.Substring(HEX_START_OF_LINE.Length);
                        }
                        currPropValue = currPropValue + valuePart;
                    }

                    else if (line.StartsWith("\"")) //Continued string literal property ***
                    {
                        currPropValue = currPropValue + line.Substring(1,line.Length - 2);
                    }

                    else if (line.Contains(PROPERTY_NAME_VALUE_DELIMITER) && (isMainDocumentSection || isInControlsSection)) //A Property name-value-pair
                    {

                        //First output the previous name
                        if (isTargetProperty) //TODO - output if we reach an end statement
                        {
                            //TODO - Add the property to the dictionary
                            //Debug.Print currentBlockName, CurrPropName, CurrPropValue
                        }

                        currPropName = line.Substring(1, line.IndexOf(PROPERTY_NAME_VALUE_DELIMITER));
                        if (isTargetProperty)
                        {
                            valuePart = (line.Substring(line.IndexOf(PROPERTY_NAME_VALUE_DELIMITER) + PROPERTY_NAME_VALUE_DELIMITER.Length)).Trim();
                            if (valuePart == "Begin") 
                            {
                                //There are binary values on the following lines
                            }
                            else
                            {
                                currPropValue = valuePart; //TODO - handle multi-line strings
                            }
                        } 
                        else {
                            valuePart = string.Empty;
                        }
                    }

                    else //Unrecognized line - Do nothing
                    { 
                    }
                }
            }
            return new List<string>();
        }

        private string ExportPath
        {
            get
            {
                var assemblyLocation = Assembly.GetAssembly(typeof(AccessApp)).Location;
                return Path.GetDirectoryName(assemblyLocation);
            }
        }

        private string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            //Access only supports Project.Procedure syntax. Error occurs if there are naming conflicts.
            // http://msdn.microsoft.com/en-us/library/office/ff193559(v=office.15).aspx
            // https://github.com/retailcoder/Rubberduck/issues/109

            var projectName = qualifiedMemberName.QualifiedModuleName.Project.Name;
            return string.Concat(projectName, ".", qualifiedMemberName.MemberName);
        }
    }
}
