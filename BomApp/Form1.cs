using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BomApp.Models;

namespace BomApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string _item = "";
            try
            {
                label1.Text = "Starting...";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    var fileInfo = new FileInfo(openFileDialog1.FileName);
                    using (var package = new ExcelPackage(fileInfo))
                    {
                        //I. Quantity

                        var worksheet = package.Workbook.Worksheets["BOM"];
                        //var rowCount = worksheet.Dimension.Rows;
                        //var colCount = worksheet.Dimension.Columns;

                        //1. Readata

                        var row = 1;
                        var data = new List<VttbModel>();
                        bool start = true;
                        while (start)
                        {
                            row++;

                            string Item = worksheet.Cells[row, 1].Value?.ToString();
                            string Title = worksheet.Cells[row, 2].Value?.ToString();
                            string Category = worksheet.Cells[row, 3].Value?.ToString();
                            string PartNumber = worksheet.Cells[row, 4].Value?.ToString();
                            string Subject = worksheet.Cells[row, 5].Value?.ToString();
                            string Manager = worksheet.Cells[row, 6].Value?.ToString();
                            string QTY = worksheet.Cells[row, 7].Value?.ToString();
                            string Material = worksheet.Cells[row, 8].Value?.ToString();
                            string Mass = worksheet.Cells[row, 9].Value?.ToString();
                            string Company = worksheet.Cells[row, 10].Value?.ToString();
                            string Status = worksheet.Cells[row, 11].Value?.ToString();

                            if (string.IsNullOrEmpty(Item)) { start = false; break; }

                            var item = new VttbModel
                            {
                                Item = Item,
                                Title = Title,
                                Category = Category,
                                PartNumber = PartNumber,
                                Subject = Subject,
                                Manager = Manager,
                                QTY = int.Parse(QTY),
                                Material = Material,
                                Mass = Mass,
                                Company = Company,
                                Status = Status,
                            };
                            data.Add(item);
                        }

                        //2. Writedata

                        worksheet.Cells[1, 12].Value = "Quantity";
                        //worksheet.Cells[1, 13].Value = "IsMaterial";
                        //worksheet.Cells[1, 14].Value = "IsWelding";

                        row = 1;
                        start = true;
                        while (start)
                        {
                            row++;
                            int CountChild = 0;
                            int? Quantity = 0;
                            int IsMaterial = 0;
                            int IsWelding = 0;

                            string Item = worksheet.Cells[row, 1].Value?.ToString();
                            string Category = worksheet.Cells[row, 3].Value?.ToString();
                            string QTY = worksheet.Cells[row, 7].Value?.ToString();

                            if (string.IsNullOrEmpty(Item)) { start = false; break; }

                            string pItem = Item;
                            if (Item.LastIndexOf('.') > 0)
                                pItem = Item.Substring(0, Item.LastIndexOf('.'));

                            if (pItem == Item)
                            {
                                Quantity = int.Parse(QTY);
                                if (Category.ToUpper() == "HÀN")
                                    IsWelding = 1;
                            }
                            else
                            {
                                var tmpItem = data.Where(x => x.Item == pItem).FirstOrDefault();

                                Quantity = int.Parse(QTY) * tmpItem?.Quantity;

                                if (Category.ToUpper() == "HÀN")
                                    IsWelding = 1;
                                else if (tmpItem?.Category.ToUpper() == "HÀN")
                                    IsWelding = 1;
                                else
                                    IsWelding = 0;
                            }

                            CountChild = data.Where(x => x.Item.StartsWith(Item + ".")).Count();
                            if (CountChild == 0)
                                IsMaterial = 1;
                            else
                                IsMaterial = 0;

                            var t = data.Where(x => x.Item == Item).FirstOrDefault();
                            t.Quantity = Quantity;
                            t.IsMaterial = IsMaterial;
                            t.IsWelding = IsWelding;

                            worksheet.Cells[row, 12].Value = Quantity;
                            //worksheet.Cells[row, 13].Value = IsMaterial;
                            //worksheet.Cells[row, 14].Value = IsWelding;
                        }

                        //II. DMVT
                        var dmvt = package.Workbook.Worksheets["DMVT"];
                        if (dmvt != null) package.Workbook.Worksheets.Delete("DMVT");
                        var wsDMVT = package.Workbook.Worksheets.Add("DMVT");
                        CreateSheetDMVT(wsDMVT, data);

                        //III. Material
                        var thvt = package.Workbook.Worksheets["THVT"];
                        if (thvt != null) package.Workbook.Worksheets.Delete("THVT");
                        var wsTHVT = package.Workbook.Worksheets.Add("THVT");
                        CreateSheetTHVT(wsTHVT, data);

                        //IV. Welding
                        var han = package.Workbook.Worksheets["HAN"];
                        if (han != null) package.Workbook.Worksheets.Delete("HAN");
                        var wsHAN = package.Workbook.Worksheets.Add("HAN");
                        CreateSheetHAN(wsHAN, data);

                        //V. DMP
                        var dmp = package.Workbook.Worksheets["DMP"];
                        if (dmp != null) package.Workbook.Worksheets.Delete("DMP");
                        var wsDMP = package.Workbook.Worksheets.Add("DMP");
                        CreateSheetDMP(wsDMP, data);

                        package.Save();
                    }

                    label1.Text = "Finished";
                }
            }
            catch (Exception ex) { label1.Text = "Error item " + _item + ": " + ex.Message; }
        }

        private void CreateSheetDMVT(ExcelWorksheet worksheet, List<VttbModel> data)
        {
            var tmpWelding = (from t in data
                              select t).ToList();

            var row = 15;
            worksheet.Cells[row, 1].Value = "STT";
            worksheet.Cells[row, 2].Value = "TÊN VẬT TƯ";
            worksheet.Cells[row, 3].Value = "MÃ VẬT TƯ";
            worksheet.Cells[row, 4].Value = "THÔNG SỐ";
            worksheet.Cells[row, 5].Value = "ĐV";
            worksheet.Cells[row, 6].Value = "SỐ LƯỢNG";
            worksheet.Cells[row, 7].Value = "TỔNG SL";
            worksheet.Cells[row, 8].Value = "VẬT LIỆU";
            worksheet.Cells[row, 9].Value = "KHỐI LƯỢNG";
            worksheet.Cells[row, 10].Value = "HÃNG SX";
            worksheet.Cells[row, 11].Value = "ĐƠN GIÁ";
            worksheet.Cells[row, 12].Value = "THÀNH TIỀN";
            worksheet.Cells[row, 13].Value = "GHI CHÚ";

            //row = 16;
            //worksheet.Cells[row, 1].Value = "Item";
            //worksheet.Cells[row, 2].Value = "Title";
            //worksheet.Cells[row, 3].Value = "PartNumber";
            //worksheet.Cells[row, 4].Value = "Category";
            //worksheet.Cells[row, 5].Value = "Manager";
            //worksheet.Cells[row, 6].Value = "QTY";
            //worksheet.Cells[row, 7].Value = "";//Quantity
            //worksheet.Cells[row, 8].Value = "Material";
            //worksheet.Cells[row, 9].Value = "Mass";
            //worksheet.Cells[row, 10].Value = "Company";

            foreach (var item in tmpWelding)
            {
                row++;
                worksheet.Cells[row, 1].Value = item.Item;
                worksheet.Cells[row, 2].Value = item.Title;
                worksheet.Cells[row, 3].Value = item.PartNumber;
                worksheet.Cells[row, 4].Value = item.Category;
                worksheet.Cells[row, 5].Value = item.Manager;
                worksheet.Cells[row, 6].Value = item.QTY;
                worksheet.Cells[row, 7].Value = item.Quantity;
                worksheet.Cells[row, 8].Value = item.Material;
                worksheet.Cells[row, 9].Value = item.Mass;
                worksheet.Cells[row, 10].Value = item.Company;
                worksheet.Cells[row, 12].Formula = $"K{row}*G{row}";
            }

            int rowSign = row + 2;
            worksheet.Cells[rowSign, 2].Value = "NGƯỜI LẬP";
            worksheet.Cells[rowSign, 11].Value = "NGƯỜI DUYỆT";

            //Style

            worksheet.Column(1).Width = 6;
            worksheet.Column(2).Width = 40;
            worksheet.Column(3).Width = 30;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 10;
            worksheet.Column(6).Width = 15;
            worksheet.Column(7).Width = 15;
            worksheet.Column(8).Width = 15;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 12;
            worksheet.Column(11).Width = 10;
            worksheet.Column(12).Width = 20;
            worksheet.Column(13).Width = 10;

            var cells = worksheet.Cells[15, 1, row, 13];
            cells.Style.Font.Name = "Times New Roman";
            cells.Style.Font.Size = 12;
            cells.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            cells = worksheet.Cells[15, 1, 15, 13];
            cells.Style.Font.Bold = true;
            cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //--Title

            var imagePath = "assets\\logo_dmp.jpg";
            var picture = worksheet.Drawings.AddPicture("MyPicture", imagePath);
            picture.SetPosition(1, 0, 0, 0); // Dòng 2 (index 1), Cột 1 (index 0)
            picture.SetSize(250, 100); // Chiều rộng 100px, chiều cao 100px

            worksheet.Cells[2, 1, 6, 13].Merge = true;
            worksheet.Cells[2, 1, 6, 13].Style.WrapText = true;
            var cell = worksheet.Cells[2, 1];
            cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cell.Style.Font.Name = "Times New Roman";
            cell.Style.Font.Size = 12;

            cell.Value = null;
            var text1 = cell.RichText.Add("CÔNG TY TNHH DMP MACHINERY");
            text1.Bold = true;
            text1.Size = 20;
            cell.RichText.Add("\n");
            var text2 = cell.RichText.Add("Địa chỉ: Số Nhà 25 Ngách 193 / 15 Phố Cầu Cốc, P.Tây Mỗ, Q. Nam Từ Liêm, TP.Hà Nội");
            text2.Bold = false;
            text2.Size = 12;
            cell.RichText.Add("\n");
            var text3 = cell.RichText.Add("Hotline: 0973992528");
            text3.Bold = false;
            text3.Size = 12;
            cell.RichText.Add("\n");
            var text4 = cell.RichText.Add("Website: www.dmpmachinery.vn - Email: info@dmpmachinery.vn");
            text4.Bold = false;
            text4.Size = 12;

            worksheet.Cells[7, 1, 8, 13].Merge = true;
            worksheet.Cells[7, 1].Value = "DANH MỤC VẬT TƯ";
            worksheet.Cells[7, 1].Style.Font.Bold = true;
            worksheet.Cells[7, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[7, 1].Style.Font.Size = 20;
            worksheet.Cells[7, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[7, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells[9, 1, 9, 13].Merge = true;
            worksheet.Cells[9, 1].Value = "Tên dự án: ";
            worksheet.Cells[9, 1].Style.Font.Bold = true;
            worksheet.Cells[9, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[9, 1].Style.Font.Size = 12;

            worksheet.Cells[10, 1, 10, 13].Merge = true;
            worksheet.Cells[10, 1].Value = "Mã dự án: ";
            worksheet.Cells[10, 1].Style.Font.Bold = true;
            worksheet.Cells[10, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[10, 1].Style.Font.Size = 12;

            worksheet.Cells[11, 1, 11, 13].Merge = true;
            worksheet.Cells[11, 1].Value = "Mã thiết kế: ";
            worksheet.Cells[11, 1].Style.Font.Bold = true;
            worksheet.Cells[11, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[11, 1].Style.Font.Size = 12;

            worksheet.Cells[12, 1, 12, 13].Merge = true;
            worksheet.Cells[12, 1].Value = "Thiết kế cơ: ";
            worksheet.Cells[12, 1].Style.Font.Bold = true;
            worksheet.Cells[12, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[12, 1].Style.Font.Size = 12;

            worksheet.Cells[13, 1, 13, 13].Merge = true;
            worksheet.Cells[13, 1].Value = "Thiết kế điện: ";
            worksheet.Cells[13, 1].Style.Font.Bold = true;
            worksheet.Cells[13, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[13, 1].Style.Font.Size = 12;

            //--Sign
            cells = worksheet.Cells[rowSign, 1, rowSign, 12];
            cells.Style.Font.Bold = true;
            cells.Style.Font.Size = 12;
            cells.Style.Font.Name = "Times New Roman";
        }

        private void CreateSheetTHVT(ExcelWorksheet worksheet, List<VttbModel> data)
        {
            var tmpMaterial = (from t1 in data
                               where t1.IsMaterial == 1 && t1.IsWelding == 0
                               group t1 by new { t1.Title, t1.PartNumber, t1.Category, t1.Company } into x1
                               select new VttbGroupModel
                               {
                                   Title = x1.Key.Title,
                                   PartNumber = x1.Key.PartNumber,
                                   Category = x1.Key.Category,
                                   Company = x1.Key.Company,
                                   Quantity = x1.Sum(z => z.Quantity)
                               })
                               .ToList();

            var row = 15;
            worksheet.Cells[row, 1].Value = "STT";
            worksheet.Cells[row, 2].Value = "TÊN VẬT TƯ";
            worksheet.Cells[row, 3].Value = "MÃ VẬT TƯ";
            worksheet.Cells[row, 4].Value = "THÔNG SỐ";
            worksheet.Cells[row, 5].Value = "ĐV";
            worksheet.Cells[row, 6].Value = "SỐ LƯỢNG";
            worksheet.Cells[row, 7].Value = "TỔNG SL";
            worksheet.Cells[row, 8].Value = "VẬT LIỆU";
            worksheet.Cells[row, 9].Value = "KHỐI LƯỢNG";
            worksheet.Cells[row, 10].Value = "HÃNG SX";
            worksheet.Cells[row, 11].Value = "ĐƠN GIÁ";
            worksheet.Cells[row, 12].Value = "THÀNH TIỀN";
            worksheet.Cells[row, 13].Value = "GHI CHÚ";

            //row = 2;
            //worksheet.Cells[row, 1].Value = "Item";
            //worksheet.Cells[row, 2].Value = "Title";
            //worksheet.Cells[row, 3].Value = "PartNumber";
            //worksheet.Cells[row, 4].Value = "Category";
            //worksheet.Cells[row, 5].Value = "Manager";
            //worksheet.Cells[row, 6].Value = "QTY";
            //worksheet.Cells[row, 7].Value = "";//Quantity
            //worksheet.Cells[row, 8].Value = "Material";
            //worksheet.Cells[row, 9].Value = "Mass";
            //worksheet.Cells[row, 10].Value = "Company";

            foreach (var item in tmpMaterial)
            {
                row++;
                worksheet.Cells[row, 2].Value = item.Title;
                worksheet.Cells[row, 3].Value = item.PartNumber;
                worksheet.Cells[row, 4].Value = item.Category;
                worksheet.Cells[row, 7].Value = item.Quantity;
                worksheet.Cells[row, 10].Value = item.Company;
            }

            int rowSign = row + 2;
            worksheet.Cells[rowSign, 2].Value = "NGƯỜI LẬP";
            worksheet.Cells[rowSign, 11].Value = "NGƯỜI DUYỆT";

            //Style

            worksheet.Column(1).Width = 6;
            worksheet.Column(2).Width = 40;
            worksheet.Column(3).Width = 30;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 10;
            worksheet.Column(6).Width = 15;
            worksheet.Column(7).Width = 15;
            worksheet.Column(8).Width = 15;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 12;
            worksheet.Column(11).Width = 10;
            worksheet.Column(12).Width = 20;
            worksheet.Column(13).Width = 10;

            var cells = worksheet.Cells[15, 1, row, 13];
            cells.Style.Font.Name = "Times New Roman";
            cells.Style.Font.Size = 12;
            cells.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            cells = worksheet.Cells[15, 1, 15, 13];
            cells.Style.Font.Bold = true;
            cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //--Title

            var imagePath = "assets\\logo_dmp.jpg";
            var picture = worksheet.Drawings.AddPicture("MyPicture", imagePath);
            picture.SetPosition(1, 0, 0, 0); // Dòng 2 (index 1), Cột 1 (index 0)
            picture.SetSize(250, 100); // Chiều rộng 100px, chiều cao 100px

            worksheet.Cells[2, 1, 6, 13].Merge = true;
            worksheet.Cells[2, 1, 6, 13].Style.WrapText = true;
            var cell = worksheet.Cells[2, 1];
            cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cell.Style.Font.Name = "Times New Roman";
            cell.Style.Font.Size = 12;

            cell.Value = null;
            var text1 = cell.RichText.Add("CÔNG TY TNHH DMP MACHINERY");
            text1.Bold = true;
            text1.Size = 20;
            cell.RichText.Add("\n");
            var text2 = cell.RichText.Add("Địa chỉ: Số Nhà 25 Ngách 193 / 15 Phố Cầu Cốc, P.Tây Mỗ, Q. Nam Từ Liêm, TP.Hà Nội");
            text2.Bold = false;
            text2.Size = 12;
            cell.RichText.Add("\n");
            var text3 = cell.RichText.Add("Hotline: 0973992528");
            text3.Bold = false;
            text3.Size = 12;
            cell.RichText.Add("\n");
            var text4 = cell.RichText.Add("Website: www.dmpmachinery.vn - Email: info@dmpmachinery.vn");
            text4.Bold = false;
            text4.Size = 12;

            worksheet.Cells[7, 1, 8, 13].Merge = true;
            worksheet.Cells[7, 1].Value = "TỔNG HỢP VẬT TƯ";
            worksheet.Cells[7, 1].Style.Font.Bold = true;
            worksheet.Cells[7, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[7, 1].Style.Font.Size = 20;
            worksheet.Cells[7, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[7, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells[9, 1, 9, 13].Merge = true;
            worksheet.Cells[9, 1].Value = "Tên dự án: ";
            worksheet.Cells[9, 1].Style.Font.Bold = true;
            worksheet.Cells[9, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[9, 1].Style.Font.Size = 12;

            worksheet.Cells[10, 1, 10, 13].Merge = true;
            worksheet.Cells[10, 1].Value = "Mã dự án: ";
            worksheet.Cells[10, 1].Style.Font.Bold = true;
            worksheet.Cells[10, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[10, 1].Style.Font.Size = 12;

            worksheet.Cells[11, 1, 11, 13].Merge = true;
            worksheet.Cells[11, 1].Value = "Mã thiết kế: ";
            worksheet.Cells[11, 1].Style.Font.Bold = true;
            worksheet.Cells[11, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[11, 1].Style.Font.Size = 12;

            worksheet.Cells[12, 1, 12, 13].Merge = true;
            worksheet.Cells[12, 1].Value = "Thiết kế cơ: ";
            worksheet.Cells[12, 1].Style.Font.Bold = true;
            worksheet.Cells[12, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[12, 1].Style.Font.Size = 12;

            worksheet.Cells[13, 1, 13, 13].Merge = true;
            worksheet.Cells[13, 1].Value = "Thiết kế điện: ";
            worksheet.Cells[13, 1].Style.Font.Bold = true;
            worksheet.Cells[13, 1].Style.Font.Name = "Times New Roman";
            worksheet.Cells[13, 1].Style.Font.Size = 12;

            //--Sign
            cells = worksheet.Cells[rowSign, 1, rowSign, 12];
            cells.Style.Font.Bold = true;
            cells.Style.Font.Size = 12;
            cells.Style.Font.Name = "Times New Roman";
        }

        private void CreateSheetHAN(ExcelWorksheet worksheet, List<VttbModel> data)
        {
            var tmpWelding = (from t in data
                              where t.IsWelding == 1
                              select t).ToList();

            var row = 1;
            worksheet.Cells[row, 1].Value = "STT";
            worksheet.Cells[row, 2].Value = "TÊN VẬT TƯ";
            worksheet.Cells[row, 3].Value = "MÃ VẬT TƯ";
            worksheet.Cells[row, 4].Value = "THÔNG SỐ";
            worksheet.Cells[row, 5].Value = "ĐV";
            worksheet.Cells[row, 6].Value = "SỐ LƯỢNG";
            worksheet.Cells[row, 7].Value = "TỔNG SL";
            worksheet.Cells[row, 8].Value = "VẬT LIỆU";
            worksheet.Cells[row, 9].Value = "KHỐI LƯỢNG";
            worksheet.Cells[row, 10].Value = "HÃNG SX";
            worksheet.Cells[row, 11].Value = "ĐƠN GIÁ";
            worksheet.Cells[row, 12].Value = "THÀNH TIỀN";
            worksheet.Cells[row, 13].Value = "GHI CHÚ";

            //row = 2;
            //worksheet.Cells[row, 1].Value = "Item";
            //worksheet.Cells[row, 2].Value = "Title";
            //worksheet.Cells[row, 3].Value = "PartNumber";
            //worksheet.Cells[row, 4].Value = "Category";
            //worksheet.Cells[row, 5].Value = "Manager";
            //worksheet.Cells[row, 6].Value = "QTY";
            //worksheet.Cells[row, 7].Value = "";//Quantity
            //worksheet.Cells[row, 8].Value = "Material";
            //worksheet.Cells[row, 9].Value = "Mass";
            //worksheet.Cells[row, 10].Value = "Company";

            foreach (var item in tmpWelding)
            {
                row++;
                worksheet.Cells[row, 1].Value = item.Item;
                worksheet.Cells[row, 2].Value = item.Title;
                worksheet.Cells[row, 3].Value = item.PartNumber;
                worksheet.Cells[row, 4].Value = item.Category;
                worksheet.Cells[row, 5].Value = item.Manager;
                worksheet.Cells[row, 6].Value = item.QTY;
                worksheet.Cells[row, 7].Value = item.Quantity;
                worksheet.Cells[row, 8].Value = item.Material;
                worksheet.Cells[row, 9].Value = item.Mass;
                worksheet.Cells[row, 10].Value = item.Company;
            }

            //Style

            worksheet.Column(1).Width = 6;
            worksheet.Column(2).Width = 40;
            worksheet.Column(3).Width = 30;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 10;
            worksheet.Column(6).Width = 15;
            worksheet.Column(7).Width = 15;
            worksheet.Column(8).Width = 15;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 12;
            worksheet.Column(11).Width = 10;
            worksheet.Column(12).Width = 20;
            worksheet.Column(13).Width = 10;

            var cells = worksheet.Cells[1, 1, row, 13];
            cells.Style.Font.Name = "Times New Roman";
            cells.Style.Font.Size = 12;
            cells.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            cells = worksheet.Cells[1, 1, 1, 13];
            cells.Style.Font.Bold = true;
            cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        }

        private void CreateSheetDMP(ExcelWorksheet worksheet, List<VttbModel> data)
        {
            var tmpWelding = (from t in data
                              where t.IsWelding == 0 && t.Company.ToUpper() == "DMP"
                              select t).ToList();

            var row = 1;
            worksheet.Cells[row, 1].Value = "STT";
            worksheet.Cells[row, 2].Value = "TÊN VẬT TƯ";
            worksheet.Cells[row, 3].Value = "MÃ VẬT TƯ";
            worksheet.Cells[row, 4].Value = "THÔNG SỐ";
            worksheet.Cells[row, 5].Value = "ĐV";
            worksheet.Cells[row, 6].Value = "SỐ LƯỢNG";
            worksheet.Cells[row, 7].Value = "TỔNG SL";
            worksheet.Cells[row, 8].Value = "VẬT LIỆU";
            worksheet.Cells[row, 9].Value = "KHỐI LƯỢNG";
            worksheet.Cells[row, 10].Value = "HÃNG SX";
            worksheet.Cells[row, 11].Value = "ĐƠN GIÁ";
            worksheet.Cells[row, 12].Value = "THÀNH TIỀN";
            worksheet.Cells[row, 13].Value = "GHI CHÚ";

            //row = 2;
            //worksheet.Cells[row, 1].Value = "Item";
            //worksheet.Cells[row, 2].Value = "Title";
            //worksheet.Cells[row, 3].Value = "PartNumber";
            //worksheet.Cells[row, 4].Value = "Category";
            //worksheet.Cells[row, 5].Value = "Manager";
            //worksheet.Cells[row, 6].Value = "QTY";
            //worksheet.Cells[row, 7].Value = "";//Quantity
            //worksheet.Cells[row, 8].Value = "Material";
            //worksheet.Cells[row, 9].Value = "Mass";
            //worksheet.Cells[row, 10].Value = "Company";

            foreach (var item in tmpWelding)
            {
                row++;
                worksheet.Cells[row, 1].Value = item.Item;
                worksheet.Cells[row, 2].Value = item.Title;
                worksheet.Cells[row, 3].Value = item.PartNumber;
                worksheet.Cells[row, 4].Value = item.Category;
                worksheet.Cells[row, 5].Value = item.Manager;
                worksheet.Cells[row, 6].Value = item.QTY;
                worksheet.Cells[row, 7].Value = item.Quantity;
                worksheet.Cells[row, 8].Value = item.Material;
                worksheet.Cells[row, 9].Value = item.Mass;
                worksheet.Cells[row, 10].Value = item.Company;
            }

            //Style

            worksheet.Column(1).Width = 6;
            worksheet.Column(2).Width = 40;
            worksheet.Column(3).Width = 30;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 10;
            worksheet.Column(6).Width = 15;
            worksheet.Column(7).Width = 15;
            worksheet.Column(8).Width = 15;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 12;
            worksheet.Column(11).Width = 10;
            worksheet.Column(12).Width = 20;
            worksheet.Column(13).Width = 10;

            var cells = worksheet.Cells[1, 1, row, 13];
            cells.Style.Font.Name = "Times New Roman";
            cells.Style.Font.Size = 12;
            cells.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            cells = worksheet.Cells[1, 1, 1, 13];
            cells.Style.Font.Bold = true;
            cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        }
    }
}