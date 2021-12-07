using System;
using System.IO;
using ExcelDataReader;

namespace pep_parser
{
    class ExcelParser
    {

        public ExcelParser(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            Console.WriteLine(reader.Read());
                        }
                    } while (reader.NextResult());

                    // The result of each spreadsheet is in result.Tables
                }
            }
        }

        public void CreatePersons()
        {
          
        }
    }
}
