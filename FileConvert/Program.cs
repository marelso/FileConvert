using System;
using FileConvert.Services;

namespace FileConvert
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path, sheet;
            if (args.Length == 0)
            {
                Console.WriteLine("Enter the path for '.xlsx' file: ");
                path = Console.ReadLine();
                while (string.IsNullOrEmpty(path))
                {
                    Console.WriteLine("The path cannot be empty.");
                    path = Console.ReadLine();                    
                }
                Console.WriteLine("Enter the worksheet name: (note: if you don't enter, the first worksheet will be considered).");
                sheet = Console.ReadLine();

                Console.WriteLine(FileConvertService.ConvertToCsv(path, sheet));
            }
            else
            {
                path = args[0];
                sheet = args.Length > 1 ? args[1] : null;

                Console.WriteLine(FileConvertService.ConvertToCsv(path, sheet));
            }
        }
    }
}
