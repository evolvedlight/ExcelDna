Installing Excel-Dna
====================

With the package manager console: ::

    Install-Package ExcelDna.AddIn

Add a basic function: ::

    using ExcelDna.Integration;

    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }
    }

Compile and run to use it in Excel: ::

    =SayHello("World!")