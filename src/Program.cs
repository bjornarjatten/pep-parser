using System;

namespace pep_parser
{
    public class Program
    {
        static string input;
        static ExcelParser parser;

        static void Main(string[] args)
        {
            // Parsing the excel sheet trough our own function (this is primarily for more clean code and making the software more testable)
            ParseExcelSheet("data/PEP_listen.xlsx", 900);
            
            // Console interface
            Console.WriteLine("Write last name...");
            string lastname = Console.ReadLine();

            Console.WriteLine("Write first name...");
            string firstname = Console.ReadLine();

            Console.WriteLine("Write birthdate name...");
            string birthdate = Console.ReadLine();

            // Writing the output of the program
            Console.WriteLine(InputConverter(lastname, firstname, birthdate));
        }

        //Splits into a seperate function to ease the test process and follow clean code principles
        public static bool InputConverter(string lastname, string firstname, string birthdate)
        {
            // Put together as one variable to follow the guidlines given by the project description
            string name = firstname + " " + lastname;

            // Return the result from searching in the parsed exceldata
            return parser.SearchPersonInExcelSheet(name, DateTime.Parse(birthdate));
        }

        public static void ParseExcelSheet(string path, int maxRows)
        {
            // This is for making sure the parser library works as accepted. 
            // For more about this bug look here: https://github.com/ExcelDataReader/ExcelDataReader
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            // Create a new instance of our own class for Excelparser
            parser = new ExcelParser(path, maxRows);
        }
    }
}
