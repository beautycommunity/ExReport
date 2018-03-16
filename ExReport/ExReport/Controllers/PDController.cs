using ClosedXML.Excel;
using ExReport.DATA;
using ExReport.Models;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExReport.Controllers
{
    public class PDController : Controller
    {
        // GET: PD
        public ActionResult Index()
        {

            return View();
        }

        public ActionResult GetPoPi(DateViewModels Data)
        {
            DataTable dte = new DataTable();

            using (DataReportDataContext db = new DataReportDataContext())
            {

                db.CommandTimeout = 60 * 10;

                //var q = db.POPI();
                //dte = q.ToDataTable();

                using (XLWorkbook wb = new XLWorkbook())
                {

                    wb.Worksheets.Add(dte, "Sheet1");
                    wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wb.Style.Font.Bold = true;

                    Response.Clear();
                    Response.Buffer = true;
                    Response.Charset = "";
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;filename= PoPi.xlsx");

                    using (MemoryStream MyMemoryStream = new MemoryStream())
                    {
                        wb.SaveAs(MyMemoryStream);
                        MyMemoryStream.WriteTo(Response.OutputStream);
                        Response.Flush();
                        Response.End();
                    }
                }
            }

            return RedirectToAction("Index", "PD");
        }
    }

}