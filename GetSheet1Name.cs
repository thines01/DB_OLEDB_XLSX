using System;
using System.Data.OleDb;
using System.Linq;

namespace DB_OLEDB_XLSX
{
   public partial class CDB_OLEDB_XLSX
   {
      public static readonly Func<OleDbConnection, string> GetSheet1Name = (conn) =>
         conn
         .GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, new[] { null, null, null, "TABLE" })
         .Select().Select(row => row["TABLE_NAME"]).FirstOrDefault().ToString();
   }
}