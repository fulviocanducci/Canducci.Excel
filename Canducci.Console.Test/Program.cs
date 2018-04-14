using System;
using Canducci.Excel;
namespace Canducci.Console.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            int[] inteiros = new int[5]
            {
                1,2,3,4,5
            };
            var headerCollection = HeaderCollection.Create();
            headerCollection.Add(Header.Create("Numeros",1));
            inteiros.ToExcelSaveAs("Inteiros.xlsx", headerCollection);

            System.Console.WriteLine("Hello World!");
        }
    }
}
