using System;
using System.Collections.Generic;
using Canducci.Excel;
namespace Canducci.Console.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //int[] inteiros = new int[5]
            //{
            //    1,2,3,4,5
            //};
            //var headerCollection = HeaderCollection.Create();
            //headerCollection.Add(Header.Create("Numeros", 1));
            //inteiros.ToExcelSaveAs("Inteiros.xlsx", headerCollection);

            IList<People> peoples = new List<People>();
            peoples.Add(new People
            {
                Id = 1,
                Name = "Name 1"
            });
            peoples.Add(new People
            {
                Id = 2,
                Name = "Name 2"
            });
            peoples.ToExcelSaveAs("Peoples.xlsx");
            var a = ContentTypeExcel.Xls.ToValue();
            var b = ContentTypeExcel.Xlsx.ToValue();
            

            System.Console.WriteLine("Hello World!");
        }
    }
}
