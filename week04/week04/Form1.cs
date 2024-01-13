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
                values[sor - 2, 0] = flat.Code;
                values[sor - 2, 1] = flat.Seller;
                values[sor - 2, 2] = flat.Side;
                values[sor - 2, 3] = flat.District;
                values[sor - 2, 4] = flat.Lift;
                values[sor - 2, 5] = flat.Rooms;
                values[sor - 2, 6] = flat.Area;
                values[sor - 2, 7] = flat.Price;
                values[sor - 2, 8] = flat.PricePerSquareMeter;
                sor++;
            }
        }
    }

}
