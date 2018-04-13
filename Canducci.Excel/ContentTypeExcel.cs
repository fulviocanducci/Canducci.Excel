using System;

namespace Canducci.Excel
{
    public static class ContentTypeExcel
    {
        private const string xls = "application/vnd.ms-excel";
        private const string xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public static String Xls => xls;

        public static String Xlsx => xlsx;
    }
}
