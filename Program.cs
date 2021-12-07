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
            parser = new ExcelParser("data/PEP_listen.xlsx");
            parser.CreatePersons();

            Console.WriteLine("Hello World!");
        }
    }
}
