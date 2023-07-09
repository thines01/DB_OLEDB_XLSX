using System.Data.OleDb;
using System.IO;

namespace DB_OLEDB_XLSX
{
   public partial class CDB_OLEDB_XLSX
   {
      public OleDbConnectionStringBuilder csb { get; set; }

      private static OleDbConnectionStringBuilder GetCsFromFi(FileInfo fi, bool blnHeader)
      {
         string strHasHeader = blnHeader ? "Yes" : "No";

         return new OleDbConnectionStringBuilder()
         {
            ConnectionString = "Extended Properties=\"Excel 12.0 Xml;HDR=" + strHasHeader + ";IMEX=1;MAXSCANROWS=2;READONLY=TRUE\";",
            DataSource = fi.FullName,
            Provider = "Microsoft.ACE.OLEDB.12.0",
         };
      }

      public CDB_OLEDB_XLSX(FileInfo fi, bool blnHeader = true)
      {
         csb = GetCsFromFi(fi, blnHeader);
      }

      public CDB_OLEDB_XLSX(string strFileName, bool blnHeader = true)
      {
         csb = GetCsFromFi(new FileInfo(strFileName), blnHeader);
      }
   }
}
