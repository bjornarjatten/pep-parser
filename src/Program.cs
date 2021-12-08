using System;

namespace pep_parser
{
    public class Program
    {
        static string input;
        static ExcelParser parser;

        static void Main(string[] args)
        {
            ParseExcelSheet("data/PEP_listen.xlsx", 900);
            Console.WriteLine("Write last name...");
            string lastname = Console.ReadLine();

            Console.WriteLine("Write first name...");
            string firstname = Console.ReadLine();

            Console.WriteLine("Write birthdate name...");
            string birthdate = Console.ReadLine();

            Console.WriteLine(InputConverter(lastname, firstname, birthdate));
        }

        //Splits into a seperate function to ease the test process
        public static bool InputConverter(string lastname, string firstname, string birthdate)
        {
            string name = firstname + " " + lastname;

            return parser.SearchPersonInExcelSheet(name, DateTime.Parse(birthdate));
        }

        public static void ParseExcelSheet(string path, int maxRows)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            parser = new ExcelParser(path, maxRows);
        }
    }
}
