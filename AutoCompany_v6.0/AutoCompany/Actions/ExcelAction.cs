using AutoCompany.Model;
using AutoCompany.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoCompany.Actions
{
    public class ExcelAction
    {
        public ExcelAction()
        {

        }
        public void CreateExcelCompany(List<LinkPage> linkPages, FormMain formMain)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileAction.AbsolutePath(@"../../Assets/Templace/Template.xlsx"))))
            {
                bool FirstAction = true;
                ExcelWorksheet workSheet;
                foreach (LinkPage linkPage in linkPages)
                {
                    int rowStart = 1;
                    package.Workbook.Worksheets.Add(linkPage.Name);
                    if (FirstAction)
                    {
                        // Lần đầu thì xóa Sheet1 mặc định đi
                        package.Workbook.Worksheets.Delete("Sheet1");
                        FirstAction = false;
                    }
                    workSheet = package.Workbook.Worksheets[linkPage.Name];
                    foreach (Companny companny in linkPage.Compannies)
                    {
                        try
                        {
                            int col = 1;
                            workSheet.Cells[rowStart, col++].Value = companny.MST;
                            workSheet.Cells[rowStart, col++].Value = companny.LicenseDate;
                            workSheet.Cells[rowStart, col++].Value = companny.Name;
                            workSheet.Cells[rowStart, col++].Value = companny.Address;
                            workSheet.Cells[rowStart, col++].Value = companny.SDT;
                            workSheet.Cells[rowStart++, col++].Value = companny.RepresentativeName;
                            App.ExcelCompany++;//THỐNG KÊ
                            if (formMain != null)
                            {
                                formMain.NotifyInfo();
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Lỗi khi ghi vào file Excel của " + linkPage.Name);
                        }
                    }
                }
                string FileWritePath = FileAction.AbsolutePath(@"../../Assets/Excel/Data.xlsx");
                Byte[] byteArrays = package.GetAsByteArray();
                File.WriteAllBytes(FileWritePath, byteArrays);
            }
        }
    }
}
