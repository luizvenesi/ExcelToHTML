using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using WebApplication4.Models;
using System.IO;
using System.Web.Hosting;

namespace WebApplication4.Controllers
{
    public class DeliveryController : Controller
    {
        
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select a excel file.<br>";
                return View("Index");
            }
            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {

                    string path = Server.MapPath("~/Excel/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);


                    //Read data from excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<Delivery> listProducts = new List<Delivery>();
                    for (int row = 3; row <= range.Rows.Count; row++)
                    {
                        Delivery p = new Delivery();
                        //p.Date = DateTime.ParseExact(((Excel.Range)range.Cells[row, 1]).Text, "dd/MM/yyyy", null);
                        p.Date = ((Excel.Range)range.Cells[row, 1]).Text;
                        p.Responsible = ((Excel.Range)range.Cells[row, 2]).Text;
                        p.Day_of_the_Week = ((Excel.Range)range.Cells[row, 3]).Text;
                        p.Driver = ((Excel.Range)range.Cells[row, 4]).Text;
                        p.Cellphone = ((Excel.Range)range.Cells[row, 5]).Text;
                        p.Company = ((Excel.Range)range.Cells[row, 6]).Text;
                        p.Order = ((Excel.Range)range.Cells[row, 7]).Text;
                        p.Type = ((Excel.Range)range.Cells[row, 8]).Text;
                        p.Obs = ((Excel.Range)range.Cells[row, 9]).Text;
                        p.Delivered = ((Excel.Range)range.Cells[row, 10]).Text;
                        p.Month = ((Excel.Range)range.Cells[row, 11]).Text;
                        p.Year = ((Excel.Range)range.Cells[row, 12]).Text;
                        p.Day = ((Excel.Range)range.Cells[row, 13]).Text;
                        p.Week = ((Excel.Range)range.Cells[row, 14]).Text;
                        listProducts.Add(p);
                    }
                    ViewBag.ListProducts = listProducts;
                    return View("Success");

                }
                else
                {
                    ViewBag.Error = "File type is incorrect <br />";
                    return View("Index");
                }
            }
        }
    }
}