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
                        worksheet.Cells[1, 13].Value = "IsMaterial";
                        //worksheet.Cells[1, 14].Value = "IsWelding";

                        row = 1;
                        start = true;
                        while (start)
                        {
                            row++;
                            int CountChild = 0;
                            int? Quantity = 0;
                            int IsMaterial = 0;
                            //int IsWelding = 0;

                            string Item = worksheet.Cells[row, 1].Value?.ToString();
                            string QTY = worksheet.Cells[row, 7].Value?.ToString();

                            if (string.IsNullOrEmpty(Item)) { start = false; break; }

                            string pItem = Item;
                            if (Item.LastIndexOf('.') > 0)
                                pItem = Item.Substring(0, Item.LastIndexOf('.'));

                            if (pItem == Item)
                                Quantity = int.Parse(QTY);
                            else
                            {
                                var tmpItem = data.Where(x => x.Item == pItem).FirstOrDefault();
                                Quantity = int.Parse(QTY) * tmpItem.Quantity;
                            }

                            CountChild = data.Where(x => x.Item.StartsWith(Item + ".")).Count();
                            if (CountChild == 0)
                                IsMaterial = 1;
                            else
                                IsMaterial = 0;

                            var t = data.Where(x => x.Item == Item).FirstOrDefault();
                            t.Quantity = Quantity;
                            t.IsMaterial = IsMaterial;

                            worksheet.Cells[row, 12].Value = Quantity;
                            worksheet.Cells[row, 13].Value = IsMaterial;
                        }

                        //II. Material
                        var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
                        //var worksheet2 = package.Workbook.Worksheets["Sheet2"];

                        var tmpMaterial = (from t1 in data
                                           where t1.IsMaterial == 1
                                           group t1 by new { t1.PartNumber, t1.Category, t1.Company } into x1
                                           select new VttbGroupModel
                                           {
                                               PartNumber = x1.Key.PartNumber,
                                               Category = x1.Key.Category,
                                               Company = x1.Key.Company,
                                               Quantity = x1.Sum(z => z.Quantity)
                                           })
                                           .ToList();

                        row = 1;
                        worksheet2.Cells[row, 1].Value = "PartNumber";
                        worksheet2.Cells[row, 2].Value = "Category";
                        worksheet2.Cells[row, 3].Value = "Company";
                        worksheet2.Cells[row, 4].Value = "Quantity";
                        foreach (var item in tmpMaterial)
                        {
                            row++;
                            worksheet2.Cells[row, 1].Value = item.PartNumber;
                            worksheet2.Cells[row, 2].Value = item.Category;
                            worksheet2.Cells[row, 3].Value = item.Company;
                            worksheet2.Cells[row, 4].Value = item.Quantity;
                        }

                        package.Save();
                    }

                    label1.Text = "Finished";
                }
            }
            catch (Exception ex) { label1.Text = ex.Message; }
        }
    }
}