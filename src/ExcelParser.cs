using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
/* An open source library under the MIT License for parsing Excel files in C# */
using ExcelDataReader;

namespace pep_parser
{
    public class ExcelParser
    {
        // Creating a list of the simple class we created, which holds the excepted properties from the Excel sheet
        List<PEPperson> persons = new List<PEPperson>();

        // The main class of the parser takes the filepath for the excel sheet and the max rows.
        // As the program is built up in a way where there are no solid paths, the excel sheet should be changeable.
        // However, due to the way we are parsing the sheet itself, this would not be possible at this point.
        public ExcelParser(string filePath, int maxRows)
        {
            // Opens the file
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // While the current row is less than max row, we will read from stream
                // This is a function from the library
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            // Checks if the current last name is null. 
                            // If so, we except that the rest of the row is null as well
                            if(reader.GetString(1) != null)
                            {
                                // Creates a new instance of our person class to hold the parsed data and places it in this object.
                                PEPperson parsedPerson = new PEPperson();
                                parsedPerson.LastName = reader.GetString(1);
                                parsedPerson.FirstName = reader.GetString(2);
                                parsedPerson.Position = reader.GetString(3);
                                
                                // As there sometimes is non-names in the last name column that will be parsed, there might be some
                                // rows where the birthdate and date added date will be null. We therefore catch these exceptions, so
                                // the program is not failing.
                                // TODO: The parsing overall could be improved further to avoid this overall, but it would still be good to check if it fails.
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
                                // The same scenario
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

        // Function for searching in the parsed data. 
        public bool SearchPersonInExcelSheet(string name, DateTime Birthday)
        {
            // As it was a requirement that the name should be one variable, this is what is passed to the function.
            // This is again split here.
            string[] nameArray = name.Split(" ");

            // A simple LINQ-query to search trough the list
            if(persons.Any(i => i.FirstName == nameArray[0] && i.LastName == nameArray[1] && i.Birthdate == Birthday))
            {
                return true;
            }
            
            return false;
        }
    }
}
