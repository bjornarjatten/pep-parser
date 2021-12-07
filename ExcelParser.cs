using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace pep_parser
{
    class ExcelParser
    {
        List<PEPperson> persons = new List<PEPperson>();

        public ExcelParser(string filePath, int maxRows)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            if(reader.GetString(1) != null)
                            {
                                PEPperson parsedPerson = new PEPperson();
                                parsedPerson.LastName = reader.GetString(1);
                                parsedPerson.FirstName = reader.GetString(2);
                                parsedPerson.Position = reader.GetString(3);
                                try
                                {
                                    parsedPerson.Birthdate = DateTime.Parse(reader.GetString(4));
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Could not parse birthdate for " + parsedPerson.LastName);
                                    // For debug
                                    //Console.WriteLine(e);
                                }
                                try
                                {
                                    parsedPerson.Added = DateTime.Parse(reader.GetString(5));
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Could not parse added date for " + parsedPerson.LastName);
                                    // For debug
                                    //Console.WriteLine(e);
                                }
                                persons.Add(parsedPerson);
                            }
                        }
                    } while (reader.RowCount < maxRows);
                }
            }
        }

        public bool SearchPersonInExcelSheet(string name, DateTime Birthday)
        {
            string[] nameArray = name.Split(" ");
            if(persons.Any(i => i.FirstName == nameArray[0] && i.LastName == nameArray[1] && i.Birthdate == Birthday))
            {
                return true;
            }
            
            return false;
        }
    }
}
