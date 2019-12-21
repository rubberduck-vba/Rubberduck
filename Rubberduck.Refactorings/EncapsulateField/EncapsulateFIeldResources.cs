using Rubberduck.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    //TODO: All the raw strings here will be assigned keys in RubberduckUI.resx 
    //just before the 'WIP' annotation is removed
    public class EncapsulateFieldResources
    {
        public static string PreviewEndOfChangesMarker
            => "'<===== No Changes below this line =====>";

        public static string DefaultPropertyParameter => "value";

        public static string DefaultStateUDTFieldName => "this";

        public static string StateUserDefinedTypeIdentifierPrefix => "T";

        public static string GroupBoxHeaderSuffix = "Encapsulation Property Name:";

        public static string Caption
            => RubberduckUI.EncapsulateField_Caption;

        public static string InstructionText  /* => RubberduckUI.EncapsulateField_InstructionText*/
            => "Select one or more fields to encapsulate.  Optionally edit property names or accept the default value(s)";

        public static string Preview
            => RubberduckUI.EncapsulateField_Preview;

        public static string TitleText
            => RubberduckUI.EncapsulateField_TitleText;

        public static string PrivateUDTPropertyText
            => "Encapsulates Each UDT Member";

        public static string Conflict => "Conflict";

        public static string Property => "Property";

        public static string Field => "Field";

        public static string Parameter => "Parameter";

        public static string NameConflictDetected => "Name Conflict Detected";
    }
}
