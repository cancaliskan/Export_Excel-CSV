using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace exportExcel_CSV.Utils
{
    public class CsvResult : ActionResult
    {
        private DataTable _dt;

        public CsvResult(DataTable dt)
        {
            _dt = dt;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            StringBuilder sb = new StringBuilder();
            HttpResponseBase response = context.HttpContext.Response;

            response.ClearContent();
            response.Clear();

            response.ContentType = "text/plain";
            response.AddHeader("content-disposition", "attachment;filename=ExportedClientList.csv");
            response.ContentType = "text/csv";

            string[] columnNames = _dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName).
                                              ToArray();
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in _dt.Rows)
            {
                string[] fields = row.ItemArray.Select(field => field.ToString()).
                                                ToArray();
                sb.AppendLine(string.Join(",", fields));
            }

            response.Write(sb.ToString());

            response.Flush();
            response.End();
        }
    }
}