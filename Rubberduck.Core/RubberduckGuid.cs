// ReSharper disable InconsistentNaming

namespace Rubberduck
{
    public static class RubberduckGuid
    {
        // Addin Guids:
        public const string ExtensionGuid = "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66";                 // shipped
        public const string DockableWindowHostGuid = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";        // shipped

        // Rubberduck API Guids:
        public const string DeclarationClassGuid = "67940D0B-081A-45BE-B0B9-CAEAFE034BC0";          // shipped prior to 2.0.14
        public const string IdentifierReferenceClassGuid = "57F78E64-8ADF-4D81-A467-A0139B877D14";  // shipped prior to 2.0.14
        public const string ParserStateClassGuid = "28754D11-10CC-45FD-9F6A-525A65412B7A";          // shipped prior to 2.0.14
        public const string IParserStateEventsGuid = "3D8EAA28-8983-44D5-83AF-2EEC4C363079";        // shipped prior to 2.0.14

        // Unit testing Guids:
        private const string UnitTestingGuidspace = "-43F0-3B33-B105-9B8188A6F040";
        public const string AssertClassGuid = "69E194DA" + UnitTestingGuidspace;                    // shipped prior to 2.0.14
        public const string IAssertGuid = "69E194DB" + UnitTestingGuidspace;                        // added for 2.0.14
        public const string PermissiveAssertClassGuid = "40F71F29-D63F-4481-8A7D-E04A4B054501";     // shipped prior to 2.0.14
        public const string FakesProviderClassGuid = "69E194DD" + UnitTestingGuidspace;             // added for 2.0.14
        public const string IFakesProviderGuid = "69E194DE" + UnitTestingGuidspace;                 // added for 2.0.14
        public const string IFakeGuid = "69E194DF" + UnitTestingGuidspace;                          // added for 2.0.14
        public const string IVerifyGuid = "69E194E0" + UnitTestingGuidspace;                        // added for 2.0.14
        public const string IStubGuid = "69E194E1" + UnitTestingGuidspace;                          // added for 2.0.14
    }
}