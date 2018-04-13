using Canducci.Excel.Interfaces;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
namespace Canducci.Excel
{
    public static class ListToExcelExtensions
    {
        #region EntityTypeConfiguration
        public static bool ToExcelSaveAs<T>(this IEnumerable items, IExcelTypeConfiguration configuration, string path)
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(items, configuration);

            excel.SaveAs(path);

            return true;
        }
        public static byte[] ToExcelByte<T>(this IEnumerable items, IExcelTypeConfiguration configuration)
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(items, configuration);

            MemoryStream stream = FabricMemoryStream.Create();

            excel.SaveAs(stream);

            return stream.ToArray();
        }
        #endregion

        #region IQueryable
        public static bool ToExcelSaveAs<T>(this IQueryable query, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat);

            excel.SaveAs(path);

            return true;
        }
        public static byte[] ToExcelByte<T>(this IQueryable query, IHeaderCollection Headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(query, Headers, dateFormat, decimalFormat);

            MemoryStream stream = FabricMemoryStream.Create();

            excel.SaveAs(stream);

            return stream.ToArray();
        }
        #endregion IQueryable

        #region IQueryableType
        public static bool ToExcelSaveAs<T>(this IQueryable<T> query, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat);

            excel.SaveAs(path);

            return true;
        }
        public static byte[] ToExcelByte<T>(this IQueryable<T> query, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(query, headers, dateFormat, decimalFormat);

            MemoryStream stream = FabricMemoryStream.Create();

            excel.SaveAs(stream);

            return stream.ToArray();
        }
        #endregion IQueryableType

        #region IEnumerable
        public static bool ToExcelSaveAs<T>(this IEnumerable items, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat);

            excel.SaveAs(path);

            return true;
        }
        public static byte[] ToExcelByte<T>(this IEnumerable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat);

            MemoryStream stream = FabricMemoryStream.Create();

            excel.SaveAs(stream);

            return stream.ToArray();
        }
        #endregion IEnumerable

        #region IEnumerableType
        public static bool ToExcelSaveAs<T>(this IEnumerable<T> items, string path, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat);

            excel.SaveAs(path);

            return true;
        }
        public static byte[] ToExcelByte<T>(this IEnumerable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            IListToExcel<T> excel = FabricListToExcel<T>.Create(items, headers, dateFormat, decimalFormat);

            MemoryStream stream = FabricMemoryStream.Create();

            excel.SaveAs(stream);

            return stream.ToArray();
        }
        #endregion IEnumerableType
    }

    internal static class FabricListToExcel<T>
    {
        #region CreateIEnumerable
        public static IListToExcel<T> Create(IEnumerable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(items, headers, dateFormat, decimalFormat);
        }
        public static IListToExcel<T> Create(IEnumerable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(items, headers, dateFormat, decimalFormat);
        }

        public static IListToExcel<T> Create(IEnumerable items, IExcelTypeConfiguration configuration)        
        {
            return new ListToExcel<T>(items, configuration);
        }
        #endregion CreateIEnumerable

        #region CreateIQueryable
        public static IListToExcel<T> Create(IQueryable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(items, headers, dateFormat, decimalFormat);
        }
        public static IListToExcel<T> Create(IQueryable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            return new ListToExcel<T>(items, headers, dateFormat, decimalFormat);
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
