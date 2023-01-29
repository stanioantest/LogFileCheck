using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace LogFileCheck
{
    public partial class Form1 : Form
    {
        List<string> numere = new List<string>();
        List<string> rangeDeCititLslUsl = new List<string>();
        List<string> rangeDeCititAV = new List<string>();
        List<string> WeldingdistanceAV = new List<string>();
        List<string> WeldingdistanceLSL = new List<string>();
        List<string> WeldingdistanceUSL = new List<string>();
        List<List<string>> listOfLslUsl = new List<List<string>>();
        List<List<string>> listOfAV = new List<List<string>>();
        int numarDeRanduri = 0;



        public Form1()
        {
            InitializeComponent();

            SetareValoriCampuriFisiere();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InitializareRange();
            // adaugarea in Lista de List a valorilor LSL si USL
            listOfLslUsl.Add(WeldingdistanceLSL);
            listOfLslUsl.Add(WeldingdistanceUSL);
            listOfAV.Add(WeldingdistanceAV);

            // citirea din fisierul excell a valorilor AV

            for (int i = 0; i < listOfAV.Count; i++)
            {
                ReadExcelFile(rangeDeCititAV[i], i);
            }
            txt_rangeWeldingDistanceValAV.Text = WeldingdistanceAV[0];

            ComparareValoriWeldingdistanceAV();

            // citirea din fisierul excell a valorilor LSL si USL         

            for (int i = 0; i < listOfLslUsl.Count; i++)
            {
                ReadExcelFile4(rangeDeCititLslUsl[i],i);
            }
            ComparareValoriWeldingdistanceLSL();
            ComparareValoriWeldingdistanceUSL();

        }
        // citirea fisierului excel pentru a afla numarul de randuri care contin date
        public int AflareRanduriDinExcell()
        {

            string filePath = txt_logfile.Text;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);

            // Aflarea numarului de randuri care au date
            numarDeRanduri = workbook.Worksheets[1].UsedRange.Rows.Count;
            workbook.Close();
            excel.Application.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);
            return numarDeRanduri;
        }
        // functia care permite citirea din fisierul excell a valorilor AV
        public void ReadExcelFile(string rangeDeCititAV,int i)
        {

            string filePath = txt_logfile.Text;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet sheet = workbook.Worksheets[1];

            //aflarea valorilor Range din fisierul excell
            Range range = sheet.Range[rangeDeCititAV];
            //parcurgerea tuturor valorilor din range
            foreach (var cell in range.Value)
            {
                listOfAV[i].Add(cell.ToString());
            }
            workbook.Close();
            excel.Application.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(excel);
        }
        // functia care permite citirea din fisierul excell a valorilor LSL si USL
        public void ReadExcelFile4(string rangeDeCitit, int i)
        {

            string filePath = txt_logfile.Text;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet sheet = workbook.Worksheets[1];

            //aflarea valorilor Range din fisierul excell
            Range range = sheet.Range[rangeDeCitit];
            //parcurgerea tuturor valorilor din range
            foreach (var cell in range.Value)
            {
                listOfLslUsl[i].Add(cell.ToString());
            }
            workbook.Close();
            excel.Application.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(excel);
        }

        //initializarea valorilor range
        public void InitializareRange()
        {
            txt_mediacoloana.Text = AflareRanduriDinExcell().ToString();

            string WeldingdistanceAVRange = Properties.Settings.Default.txt_rangeWeldingDistance + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingDistance.Text = Properties.Settings.Default.txt_rangeWeldingDistance + AflareRanduriDinExcell().ToString();
            string WeldingdistanceLSLRange = Properties.Settings.Default.txt_rangeWeldingDistanceLSL + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingDistanceLSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceLSL + AflareRanduriDinExcell().ToString();
            string WeldingdistanceUSLRange = Properties.Settings.Default.txt_rangeWeldingDistanceUSL + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingDistanceUSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceUSL + AflareRanduriDinExcell().ToString();
            rangeDeCititAV.Add(WeldingdistanceAVRange);
            rangeDeCititLslUsl.Add(WeldingdistanceLSLRange);
            rangeDeCititLslUsl.Add(WeldingdistanceUSLRange);




        }
        public void SetareValoriCampuriFisiere()
        {
            txt_logfile.Text = Properties.Settings.Default.txt_logfile;
            txt_rangeWeldingDistanceValLSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceValLSL.ToString();
            txt_rangeWeldingDistanceValUSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceValUSL.ToString();


        }
        public void ComparareValoriWeldingdistanceAV()
        {
            for (int i = 0; i < WeldingdistanceAV.Count; i++)
            {
                if (Convert.ToDouble(WeldingdistanceAV[i]) >= Convert.ToDouble(txt_rangeWeldingDistanceValLSL.Text.ToString()) && Convert.ToDouble(WeldingdistanceAV[i]) <= Convert.ToDouble(txt_rangeWeldingDistanceValUSL.Text.ToString()))
                {
                    lbl_statusWeldingDistanceAV.Text = "OK";
                    lbl_statusWeldingDistanceAV.BackColor = Color.GreenYellow;
                }

                else
                {
                    lbl_statusWeldingDistanceAV.Text = "NOK";
                    lbl_statusWeldingDistanceAV.BackColor = Color.Red;
                    break;
                }
            }

        }
        public void ComparareValoriWeldingdistanceLSL()
        {
            for (int i = 0; i < WeldingdistanceLSL.Count; i++)
            {
                if (WeldingdistanceLSL[i].Equals(txt_rangeWeldingDistanceValLSL.Text.ToString()))
                {
                    lbl_statusWeldingDistanceLSL.Text = "OK";
                    lbl_statusWeldingDistanceLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusWeldingDistanceLSL.Text = "NOK";
                    lbl_statusWeldingDistanceLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriWeldingdistanceUSL()
        {
            for (int i = 0; i < WeldingdistanceUSL.Count; i++)
            {
                if (WeldingdistanceUSL[i].Equals(txt_rangeWeldingDistanceValUSL.Text.ToString()))
                {
                    lbl_statusWeldingDistanceUSL.Text = "OK";
                    lbl_statusWeldingDistanceUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusWeldingDistanceUSL.Text = "NOK";
                    lbl_statusWeldingDistanceUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

    }
}
