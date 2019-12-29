using Canducci.Excel.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
namespace Canducci.Excel.Internals
{
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

      public static IListToExcel<T> Create(IEnumerable items, HeaderCollection headers = null, string dateFormat = "yyyy-MM-dd", string decimalFormat = "#,##0.00")
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

      public static IListToExcel<T> Create(IEnumerable<T> items, HeaderCollection headers = null, string dateFormat = "yyyy-MM-dd", string decimalFormat = "#,##0.00")
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
      public static IListToExcel<T> Create(IQueryable query, HeaderCollection headers = null, string dateFormat = "yyyy-MM-dd", string decimalFormat = "#,##0.00")
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
      public static IListToExcel<T> Create(IQueryable<T> query, HeaderCollection headers = null, string dateFormat = "yyyy-MM-dd", string decimalFormat = "#,##0.00")
      {
         return new ListToExcel<T>(query, headers, dateFormat, decimalFormat);
      }
      #endregion CreateIQueryable
   }
}
