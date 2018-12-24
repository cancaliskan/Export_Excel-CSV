using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace exportExcel_CSV.Utils
{
    public class ExcelResult : ActionResult
    {
        private DataTable _dt;

        public ExcelResult(DataTable dt)
        {
            _dt = dt;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            HttpResponseBase response = context.HttpContext.Response;
            var grid = new GridView();

            grid.DataSource = _dt;
            grid.DataBind();

            response.Clear();
            response.ClearHeaders();
            response.ClearContent();

            response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1254");
            response.Charset = "windows-1254";
            response.AddHeader("content-disposition", "attachment;filename=ExportedClientList.xls");
            response.ContentType = "application/vnd.ms-excel";

            StringWriter sw = new StringWriter();
            HtmlTextWriter htmlTextWriter = new HtmlTextWriter(sw);

            grid.RenderControl(htmlTextWriter);

            response.Write("<meta http-equiv='Content-Type' content ='text/html; charset = 'windows-1254'/>" + sw.ToString());

            response.End();
        }
    }
}