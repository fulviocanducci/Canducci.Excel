using Canducci.Excel.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
namespace Canducci.Excel
{
    public static class ListToExcelExtensions
    {
        #region SaveAs

        private static bool SaveAs<T>(IListToExcel<T> excel, string path)
        {
            excel.SaveAs(path);
            return true;
        }

        public static bool ToExcelSaveAs<T>(this IEnumerable items, string path, Action<IExcelTypeConfiguration> configuration)
        {            
            return SaveAs(FabricListToExcel<T>.Create(items, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IEnumerable items, string path, IExcelTypeConfiguration configuration)
        {
            return SaveAs(FabricListToExcel<T>.Create(items, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IEnumerable items, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {            
            return SaveAs(FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat), path);
        }

        public static bool ToExcelSaveAs<T>(this IEnumerable<T> items, string path, Action<IExcelTypeConfiguration> configuration)
        {            
            return SaveAs(FabricListToExcel<T>.Create(items, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IEnumerable<T> items, string path, IExcelTypeConfiguration configuration)
        {
            return SaveAs(FabricListToExcel<T>.Create(items, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IEnumerable<T> items, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return SaveAs(FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat), path);
        }

        public static bool ToExcelSaveAs<T>(this IQueryable query, string path, Action<IExcelTypeConfiguration> configuration)
        {
            return SaveAs(FabricListToExcel<T>.Create(query, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IQueryable query, string path, IExcelTypeConfiguration configuration)
        {
            return SaveAs(FabricListToExcel<T>.Create(query, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IQueryable query, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return SaveAs(FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat), path);
        }

        public static bool ToExcelSaveAs<T>(this IQueryable<T> query, string path, Action<IExcelTypeConfiguration> configuration)
        {
            return SaveAs(FabricListToExcel<T>.Create(query, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IQueryable<T> query, string path, IExcelTypeConfiguration configuration)
        {
            return SaveAs(FabricListToExcel<T>.Create(query, configuration), path);
        }
        public static bool ToExcelSaveAs<T>(this IQueryable<T> query, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return SaveAs(FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat), path);
        }
        
        #endregion SaveAs

        #region ExcelByte
        
        private static byte[] ExcelByte<T>(IListToExcel<T> excel)
        {
            MemoryStream stream = FabricMemoryStream.Create();
            excel.SaveAs(stream);
            return stream.ToArray();
        }

        public static byte[] ToExcelByte<T>(this IEnumerable items, IExcelTypeConfiguration configuration)
        {            
            return ExcelByte(FabricListToExcel<T>.Create(items, configuration));
        }        
        public static byte[] ToExcelByte<T>(this IEnumerable items, Action<IExcelTypeConfiguration> configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(items, configuration));
        }
        public static byte[] ToExcelByte<T>(this IEnumerable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {            
            return ExcelByte(FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat));
        }

        public static byte[] ToExcelByte<T>(this IEnumerable<T> items, Action<IExcelTypeConfiguration> configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(items, configuration));
        }
        public static byte[] ToExcelByte<T>(this IEnumerable<T> items, IExcelTypeConfiguration configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(items, configuration));
        }
        public static byte[] ToExcelByte<T>(this IEnumerable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return ExcelByte(FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat));
        }

        public static byte[] ToExcelByte<T>(this IQueryable query, Action<IExcelTypeConfiguration> configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(query, configuration));
        }
        public static byte[] ToExcelByte<T>(this IQueryable query, IExcelTypeConfiguration configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(query, configuration));
        }
        public static byte[] ToExcelByte<T>(this IQueryable query, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return ExcelByte(FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat));
        }

        public static byte[] ToExcelByte<T>(this IQueryable<T> query, Action<IExcelTypeConfiguration> configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(query, configuration));
        }
        public static byte[] ToExcelByte<T>(this IQueryable<T> query, IExcelTypeConfiguration configuration)
        {
            return ExcelByte(FabricListToExcel<T>.Create(query, configuration));
        }
        public static byte[] ToExcelByte<T>(this IQueryable<T> query, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {            
            return ExcelByte(FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat));
        }

        #endregion ExcelByte

        #region ToValue
        public static string ToValue(this ContentTypeExcel value)           
        {
            ContentTypeAttribute result = (value
                .GetType()
                .GetField(Enum.GetName(typeof(ContentTypeExcel), value))
                .GetCustomAttributes(true)
                .FirstOrDefault()) as ContentTypeAttribute;
            if (result != null)
            {
                return result.Value;
            }                
            return default(string);
        }
        #endregion ToValue
    }

    internal static class FabricListToExcel<T>
    {
        #region CreateIEnumerable
        public static IListToExcel<T> Create(IEnumerable items, IExcelTypeConfiguration configuration)
        {
            return new ListToExcel<T>(items, configuration);
        }

        public static IListToExcel<T> Create(IEnumerable items, Action<IExcelTypeConfiguration> configuration)
        {
            return new ListToExcel<T>(items, configuration);
        }

        public static IListToExcel<T> Create(IEnumerable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(items, headers, dateFormat, decimalFormat);
        }

        public static IListToExcel<T> Create(IEnumerable<T> items, IExcelTypeConfiguration configuration)
        {
            return new ListToExcel<T>(items, configuration);
        }

        public static IListToExcel<T> Create(IEnumerable<T> items, Action<IExcelTypeConfiguration> configuration)
        {
            return new ListToExcel<T>(items, configuration);
        }

        public static IListToExcel<T> Create(IEnumerable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(items, headers, dateFormat, decimalFormat);
        }
        
        #endregion CreateIEnumerable

        #region CreateIQueryable
        public static IListToExcel<T> Create(IQueryable query, Action<IExcelTypeConfiguration> configuration)
        {
            return new ListToExcel<T>(query, configuration);
        }
        public static IListToExcel<T> Create(IQueryable query, IExcelTypeConfiguration configuration)
        {
            return new ListToExcel<T>(query, configuration);
        }
        public static IListToExcel<T> Create(IQueryable query, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(query, headers, dateFormat, decimalFormat);
        }

        public static IListToExcel<T> Create(IQueryable<T> query, Action<IExcelTypeConfiguration> configuration)
        {
            return new ListToExcel<T>(query, configuration);
        }
        public static IListToExcel<T> Create(IQueryable<T> query, IExcelTypeConfiguration configuration)
        {
            return new ListToExcel<T>(query, configuration);
        }
        public static IListToExcel<T> Create(IQueryable<T> query, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(query, headers, dateFormat, decimalFormat);
        }
        #endregion CreateIQueryable
    }

    internal static class FabricMemoryStream
    {
        public static MemoryStream Create()
        {
            return new MemoryStream();
        }
    }
}
