using System;

namespace pep_parser
{
    // Just a simple class holding the given properties from the excelsheet
    class PEPperson
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Position { get; set; }
        public DateTime Birthdate { get; set; }
        public DateTime Added { get; set; }
    }
}