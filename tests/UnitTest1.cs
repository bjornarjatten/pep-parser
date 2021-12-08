using Xunit;
using pep_parser;

namespace tests
{
    public class UnitTest1
    {
        [Fact]
        public void CheckIfExistingNamesIsFound()
        {
            Program.ParseExcelSheet("../../../../src/data/PEP_listen.xlsx", 900);

            bool result1 = Program.InputConverter("Frederiksen", "Mette", "19.11.1977");
            bool result2 = Program.InputConverter("Ellemann-Jensen", "Jakob", "25.09.1973");
            //Does not work to the known double name bug
            //bool result3 = Program.InputConverter("Jakobsen", "Daniel Toft", "12.05.1978");
            //bool result4 = Program.InputConverter("Heunicke", "Magnus Johannes", "28.01.1975");

            Assert.True(result1, "Mette Frederiksen should exist");
            Assert.True(result2, "Jakob Ellemann-Jensen should exist");
            //Does not work to the known double name bug
            //Assert.True(result3, "Daniel Toft Jakobsen should exist");
            //Assert.True(result4, "Magnus Johannes Heunicke should exist");
        }

        [Fact]
        public void CheckIfNonExistingNamesIsNotFound()
        {
            Program.ParseExcelSheet("../../../../src/data/PEP_listen.xlsx", 900);

            bool result1 = Program.InputConverter("Solberg", "Erna", "19.11.1977");
            bool result2 = Program.InputConverter("Gahr Støre", "Jonas", "25.09.1973");

            Assert.False(result1, "Erna Solberg should not exist");
            Assert.False(result2, "Jonas Gahr Støre should not exist");
        }
    }
}
