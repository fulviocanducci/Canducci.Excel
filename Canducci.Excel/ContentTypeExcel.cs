namespace Canducci.Excel
{
   public enum ContentTypeExcel
   {
      [ContentType("application/vnd.ms-excel")]
      Xls,
      [ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
      Xlsx
   }
}
