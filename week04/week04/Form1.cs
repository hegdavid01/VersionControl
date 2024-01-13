using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.Entity.Migrations.Model;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;

namespace week04
{
    public partial class Form1 : Form
    {       
        RealEstateEntities context = new RealEstateEntities();
        List<Flat> Flats;

        Excel.Application xlApp; 
        Excel.Workbook xlWB; 
        Excel.Worksheet xlSheet; 

        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();
            CreateTable();
        }
    
        private void LoadData()
        {
            Flats = Context.Flats.ToList();
        }

        private void CreateExcel() 
        {
            try
            {
                xlApp = new Excel.Application();

                xlWB = xlApp.Workbooks.Add(Missing.Value);

                xlSheet = xlWB.ActiveSheet;

                CreateTable(); 

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            string[] headers = new string[]
            {
               "Kód",
               "Eladó",
               "Oldal",
               "Kerület",
               "Lift",
               "Szobák száma",
               "Alapterület (m2)",
               "Ár (mFt)",
               "Négyzetméter ár (Ft/m2)"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, 1] = headers[0];
            }

            object[,] values = new object[Flats.Count, headers.Length];

            int sor = 0;
            foreach (Flat f in Flats)
            {
                values[sor - 2, 0] = f.Code;
                values[sor - 2, 1] = f.Seller;
                values[sor - 2, 2] = f.Side;
                values[sor - 2, 3] = f.District;
                values[sor - 2, 4] = f.Lift;
                values[sor - 2, 5] = f.Rooms;
                values[sor - 2, 6] = f.Area;
                values[sor - 2, 7] = f.Price;
                values[sor - 2, 8] = f.PricePerSquareMeter;
                sor++;
            }

            string GetCell(int x, int y)
            {
                string ExcelCoordinate = "";
                int dividend = y;
                int modulo;

                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                    dividend = (int)((dividend - modulo) / 26);
                }
                ExcelCoordinate += x.ToString();

                return ExcelCoordinate;
            }

            xlSheet.get_Range(
            GetCell(2, 1),
            GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;

            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
        }
    }

}
