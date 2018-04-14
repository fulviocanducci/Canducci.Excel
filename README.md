## Canducci EXCEL 

[![Canducci Excel](http://i666.photobucket.com/albums/vv25/netdragoon/1446163380_excel_zps5lhqezet.png)](https://www.nuget.org/packages/Canducci.Excel/)

[![NuGet](https://img.shields.io/nuget/v/Canducci.Excel.svg?style=plastic&label=version)](https://www.nuget.org/packages/Canducci.Excel/)
[![NuGet](https://img.shields.io/nuget/dt/Canducci.Excel.svg)](https://www.nuget.org/packages/Canducci.Excel/)

### Classes

__1) Interfaces__

- `IHeader`
- `IHeaderCollection`
- `IListToExcel`

__2) Class__

- `Header`
- `HeaderCollection`
- `ListToExcel`

__3) Extension method to__ `IEnumerable`, `IEnumerable<T>`, `IQueryable` e `IQueryable<T>`

- `bool ToExcelSaveAs<T>`
- `byte[] ToExcelByte<T>`

## Package Installation (NUGET)

```Csharp

PM> Install-Package Canducci.Excel

```

## How to use?

Declare o namespace `using Canducci.Excel;` 

## Array

```Csharp
//Setting the header
HeaderCollection headerData = new HeaderCollection();
headerData.Add(new Header("DateCreated", 1));

//Array
var n = new DateTime[] { DateTime.Now.Date.AddDays(-15), DateTime.Now.Date.AddDays(-16) };

//Using the extensive method to generate the file and write to disk.
n.ToExcelSaveAs("Result\\DatasSimples.xlsx", headerData);

```

## Array two-dimensional

```Csharp

int[,] c = new int[2,2];
c[0, 0] = 1; c[0, 1] = 2;
c[1, 0] = 3; c[1, 1] = 4;

//Setting the header
IHeaderCollection headersArrayBi = new HeaderCollection();
headersArrayBi.Add(2, "Col-");

//Using the extensive method to generate the file and write to disk.
c.ToExcelSaveAs<int[,]>("Result\\ArrayBi2Dimensioanl.xlsx", headersArrayBi);

```

## List

```Csharp

//Setting the header
HeaderCollection headers = new HeaderCollection();                  
headers.Add(new Header("Id", 1));
headers.Add(new Header("Nome", 2));
headers.Add(new Header("Valor", 3));
headers.Add(new Header("Data de Validade", 4));
headers.Add(new Header("Hora Stamp", 5));

//Manually creating a typed list
IList<Test> tests = new List<Test>();
tests.Add(new Test() { Id = 1, Nome = "Test 1", Valor = 10M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 2, Nome = "Test 2", Valor = 2000.69M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 3, Nome = "Test 3", Valor = 10M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 4, Nome = "Test 4", Valor = 20M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 5, Nome = "Test 5", Valor = 10M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 6, Nome = "Test 6", Valor = 2000.69M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 7, Nome = "Test 7", Valor = 10M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });
tests.Add(new Test() { Id = 8, Nome = "Test 8", Valor = 20M, 
            Data = DateTime.Now.Date, Time = DateTime.Now.TimeOfDay });

//Usando o método extensivo para gerar o arquivo e gravar em disco.
tests.ToExcelSaveAs("Result\\List.xlsx", headers);

```

## Entity Framework (IQueryable e IEnumerable)

__Simples Console ou Desktop__

```Csharp
using (AdventureWorks2014Entities db = new AdventureWorks2014Entities())
{

    //IQueryable<Person>   
    var Query = db.Person
            .AsNoTracking()
            .AsQueryable();
    
    //Setting the header
    IHeaderCollection headerEF = new HeaderCollection();
    headerEF.Add(new Header("Primeiro Nome", 1));

    //Using the extensive method to generate the file and write to disk.    
    Query.Select(x => x.FirstName)                    
        .Take(10)
        .ToExcelSaveAs("Result\\Forma1.xlsx", headerEF);

}

```

__ASP.NET MVC__

```Csharp

[Route("excel")]
public FileContentResult CreateExcel()
{

    //Array de bytes
    byte[] fileByte = null;

    using (AdventureWorks2014Entities db = new AdventureWorks2014Entities())
    {

        //Setting the header
        IHeaderCollection headers = HeaderCollection.Create();
        headers.Add(Header.Create("Id", 1));
        headers.Add(Header.Create("Sexo", 2));
        headers.Add(Header.Create("Data", 3));
        headers.Add(Header.Create("Trabalho", 4));

        //Using the Extensive Method to Generate an Array of Bytes in Memory
        //by the extensive method ToExcelByte
        fileByte = db.Employee
            .AsNoTracking()
            .Select(x => new
            {
                x.BusinessEntityID,
                x.Gender,
                x.HireDate,
                x.JobCandidate
            })                    
            .Take(20)
            .ToExcelByte(headers);

    }           
    
    //The browser will download the file
    return File(fileByte, ContentTypeExcel.Xlsx.ToValue(), "file.xlsx");

}

```

__ASP.NET MVC__ `ListToExcel`

```Csharp
[Route("excel")]
public FileContentResult CreateExcel1()
{
    //Array of Bytes
    byte[] fileByte = null;

    //Entity Framework and MemoryStream
    using (AdventureWorks2014Entities db = new AdventureWorks2014Entities())
    using (System.IO.MemoryStream Stream = new System.IO.MemoryStream())            
    {
        //Setting the header
        IHeaderCollection Headers = HeaderCollection.Create();
        Headers.Add(Header.Create("Departamento", 1));
        Headers.Add(Header.Create("Razão Social", 2));

        //Building Query and Returning an IQueryable
        IQueryable Query = db.Department.Select(x => new { x.DepartmentID, x.Name }).AsQueryable();

        //Instantiating the ListToExcel class
        IListToExcel<Department> listDep = new ListToExcel<Department>(Query, Headers);

        //SaveAs passing values ​​to the MemoryStream (`Stream`)
        listDep.SaveAs(Stream);

        //Assigning values ​​to Array of Bytes
        fileByte = Stream.ToArray();

    }

    //The browser will download the file
    return File(fileByte, ContentTypeExcel.Xlsx.ToValue(), "file.xlsx");

}
```

__ASP.NET WebForms__

```Csharp
protected void BtnEnviar_Click(object sender, EventArgs e)
{
    var Items = new[]
    {
        new {Id  = 1, Name = "Id 1", Value = 150.00M},
        new {Id  = 2, Name = "Id 2", Value = 250.25M},
        new {Id  = 3, Name = "Id 3", Value = 450.00M}
    }.ToArray();

    byte[] ArrayOfBytes = Items.ToExcelByte();


    Response.ClearContent();
    Response.ClearHeaders();
    

    Response.AddHeader("content-disposition", "attachment; filename=file.xls");

    Response.ContentType = "application/ms-excel";

    Response.BinaryWrite(ArrayOfBytes);

    Response.End();
    
}

```

#### Notes:

It is not mandatory to use the class `HeaderCollection`, but it is a way of setting the title and order of each column. The order of values ​​must also follow the order of what was placed in each title always starting from number 1 (ex. 1,2,3, being less than or equal to 0 (zero) causes an Exception).

___Example:___

```csharp
IHeaderCollection Headers = HeaderCollection.Create();
Headers.Add(Header.Create("Departamento", 1));
Headers.Add(Header.Create("Razão Social", 2));
```

In the example above, two columns were created starting with the Department and the next Social Title.

- The file generated in Excel has the purpose of simple transport of each information, taking its type from each line and column corresponding to what was sent. 

- The package has no layout formatting (although the title is centralized by default and the columns follow the formatting according to the type sent), fonts, colors, etc., the concern is just sending an information to be modified in an excel file.

- The generated file is 100% compatible with ___Microsoft Office Excel___

### Project Test

- https://github.com/fulviocanducci/Canducci.Excel/tree/master/Canducci.Web.Test45
- https://github.com/fulviocanducci/Canducci.Excel/tree/master/Canducci.Web.Test46
- https://github.com/fulviocanducci/Canducci.Excel/tree/master/Canducci.Web.TestNETStandard2.0
