using System;

namespace pep_parser
{
    class Program
    {
        static string input;
        static ExcelParser parser;

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            parser = new ExcelParser("data/PEP_listen.xlsx", 900);

            Console.WriteLine("Write last name...");
            string lastname = Console.ReadLine();

            Console.WriteLine("Write first name...");
            string firstname = Console.ReadLine();

            Console.WriteLine("Write birthdate name...");
            string birthdate = Console.ReadLine();
            
            string name = firstname + " " + lastname;

            Console.WriteLine(parser.SearchPersonInExcelSheet(name, DateTime.Parse(birthdate)));
        }
    }
}
