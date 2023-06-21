using ExcelDna.Integration;

namespace ExcelFunctions
{
    public class Class1
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

    }
}