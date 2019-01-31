using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Data;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Import_Export_Excel.Models;

namespace Import_Export_Excel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.StudentList = null;
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelFile)
        {
            try
            {
                if (excelFile == null || excelFile.ContentLength == 0)
                {
                    ViewBag.Error = "Please select a excel file";
                    return View("Index");
                }
                else
                {
                    ViewBag.Error = "0";
                    if (excelFile.FileName.EndsWith("xls") || excelFile.FileName.EndsWith("xlsx"))
                    {
                        string path = Server.MapPath("~/FileExcel/" + excelFile.FileName);
                        //            string path2 = path.Replace("\\", "/");
                        //            string[] listFileName = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories)
                        //.Where(x => x.EndsWith(".xls") || x.EndsWith(".xlsx")).Select(Path.GetFileName).ToArray();
                        //            var s = listFileName.Where(c => c.Equals(excelFile.FileName)).FirstOrDefault();
                        //            if (s != null)
                        //                System.IO.File.Delete(path);

                        if (System.IO.File.Exists(path))
                        {
                            foreach (var process in Process.GetProcessesByName("Microsoft Office Excel (32bit)"))
                            {
                                var pro = process;
                             //   process.Kill();
                            }
                            System.IO.File.Delete(path);
                        }

                        //save file to path by path
                        excelFile.SaveAs(path);
                        //Read data from excel file
                        Excel.Application application = new Excel.Application();
                        Excel.Workbook workbook = application.Workbooks.Open(path);
                        Excel.Worksheet worksheet = workbook.ActiveSheet;
                        Excel.Range range = worksheet.UsedRange;
                        List<Student_Import> studen_list = new List<Student_Import>();
                        for (int rows = 1; rows < range.Rows.Count; rows++)
                        {
                            Student_Import si = new Student_Import
                            {
                                Student_ID = ((Excel.Range)range.Cells[rows, 1]).Text,
                                Student_Name = ((Excel.Range)range.Cells[rows, 2]).Text,
                            };
                            studen_list.Add(si);
                        }

                        ViewBag.StudentList = studen_list;

                        return View("Index");
                    }
                    else
                    {
                        ViewBag.Error = "File Type is incorrect";
                        return View("Index");
                    }
                }
            }
            catch (Exception ex)
            {
                ViewBag.Error = "Error Exception is : " + ex.Message;
                return View("Index");
            }
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}