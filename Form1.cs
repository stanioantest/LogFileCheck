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
using Label = System.Windows.Forms.Label;
using System.Threading;

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

        List<string> WeldingEnergyAV = new List<string>();
        List<string> WeldingEnergyLSL = new List<string>();
        List<string> WeldingEnergyUSL = new List<string>();

        List<string> AirFlowFlowVolumeAV = new List<string>();
        List<string> AirFlowFlowVolumeLSL = new List<string>();
        List<string> AirFlowFlowVolumeUSL = new List<string>();

        List<string> Scara1FittingPins0AV = new List<string>();
        List<string> Scara1FittingPins0LSL = new List<string>();
        List<string> Scara1FittingPins0USL = new List<string>();

        List<string> Scara1FittingPins1AV = new List<string>();
        List<string> Scara1FittingPins1LSL = new List<string>();
        List<string> Scara1FittingPins1USL = new List<string>();

        List<string> Scara1FittingPins2AV = new List<string>();
        List<string> Scara1FittingPins2LSL = new List<string>();
        List<string> Scara1FittingPins2USL = new List<string>();

        List<string> Scara1FittingPins3AV = new List<string>();
        List<string> Scara1FittingPins3LSL = new List<string>();
        List<string> Scara1FittingPins3USL = new List<string>();

        List<string> Scara1FittingPins4AV = new List<string>();
        List<string> Scara1FittingPins4LSL = new List<string>();
        List<string> Scara1FittingPins4USL = new List<string>();

        List<string> Scara2FittingPins0AV = new List<string>();
        List<string> Scara2FittingPins0LSL = new List<string>();
        List<string> Scara2FittingPins0USL = new List<string>();

        List<string> Scara2FittingPins1AV = new List<string>();
        List<string> Scara2FittingPins1LSL = new List<string>();
        List<string> Scara2FittingPins1USL = new List<string>();

        List<string> Scara2FittingPins2AV = new List<string>();
        List<string> Scara2FittingPins2LSL = new List<string>();
        List<string> Scara2FittingPins2USL = new List<string>();

        List<string> Scara2FittingPins3AV = new List<string>();
        List<string> Scara2FittingPins3LSL = new List<string>();
        List<string> Scara2FittingPins3USL = new List<string>();

        List<string> Scara2FittingPins4AV = new List<string>();
        List<string> Scara2FittingPins4LSL = new List<string>();
        List<string> Scara2FittingPins4USL = new List<string>();


        List<string> Scara3FittingPins0AV = new List<string>();
        List<string> Scara3FittingPins0LSL = new List<string>();
        List<string> Scara3FittingPins0USL = new List<string>();

        List<string> Scara3FittingPins1AV = new List<string>();
        List<string> Scara3FittingPins1LSL = new List<string>();
        List<string> Scara3FittingPins1USL = new List<string>();

        List<string> Scara3FittingPins2AV = new List<string>();
        List<string> Scara3FittingPins2LSL = new List<string>();
        List<string> Scara3FittingPins2USL = new List<string>();

        List<string> Scara3FittingPins3AV = new List<string>();
        List<string> Scara3FittingPins3LSL = new List<string>();
        List<string> Scara3FittingPins3USL = new List<string>();

        List<string> Scara3FittingPins4AV = new List<string>();
        List<string> Scara3FittingPins4LSL = new List<string>();
        List<string> Scara3FittingPins4USL = new List<string>();


        List<List<string>> listOfLslUsl = new List<List<string>>();
        List<List<string>> listOfAV = new List<List<string>>();
        List<string> txt_rangeValLSL = new List<string>();
        List<string> txt_rangeValUSL = new List<string>();
        List<Label> lbl_status = new List<Label>();
        List<Label> lbl_statusBackColor = new List<Label>();

        int numarDeRanduri = 0;



        public Form1()
        {
            InitializeComponent();

            SetareValoriCampuriFisiere();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           

                InitializareRange();
            // adaugarea in Lista de List a valorilor LSL - USL si AV
            listOfLslUsl.Add(WeldingdistanceLSL);
            listOfLslUsl.Add(WeldingdistanceUSL);
            listOfAV.Add(WeldingdistanceAV);

            listOfLslUsl.Add(WeldingEnergyLSL);
            listOfLslUsl.Add(WeldingEnergyUSL);
            listOfAV.Add(WeldingEnergyAV);

            listOfLslUsl.Add(AirFlowFlowVolumeLSL);
            listOfLslUsl.Add(AirFlowFlowVolumeUSL);
            listOfAV.Add(AirFlowFlowVolumeAV);

            listOfLslUsl.Add(Scara1FittingPins0LSL);
            listOfLslUsl.Add(Scara1FittingPins0USL);
            listOfAV.Add(Scara1FittingPins0AV);

            listOfLslUsl.Add(Scara1FittingPins1LSL);
            listOfLslUsl.Add(Scara1FittingPins1USL);
            listOfAV.Add(Scara1FittingPins1AV);

            listOfLslUsl.Add(Scara1FittingPins2LSL);
            listOfLslUsl.Add(Scara1FittingPins2USL);
            listOfAV.Add(Scara1FittingPins2AV);

            listOfLslUsl.Add(Scara1FittingPins3LSL);
            listOfLslUsl.Add(Scara1FittingPins3USL);
            listOfAV.Add(Scara1FittingPins3AV);

            listOfLslUsl.Add(Scara1FittingPins4LSL);
            listOfLslUsl.Add(Scara1FittingPins4USL);
            listOfAV.Add(Scara1FittingPins4AV);

            listOfLslUsl.Add(Scara2FittingPins0LSL);
            listOfLslUsl.Add(Scara2FittingPins0USL);
            listOfAV.Add(Scara2FittingPins0AV);

            listOfLslUsl.Add(Scara2FittingPins1LSL);
            listOfLslUsl.Add(Scara2FittingPins1USL);
            listOfAV.Add(Scara2FittingPins1AV);

            listOfLslUsl.Add(Scara2FittingPins2LSL);
            listOfLslUsl.Add(Scara2FittingPins2USL);
            listOfAV.Add(Scara2FittingPins2AV);

            listOfLslUsl.Add(Scara2FittingPins3LSL);
            listOfLslUsl.Add(Scara2FittingPins3USL);
            listOfAV.Add(Scara2FittingPins3AV);

            listOfLslUsl.Add(Scara2FittingPins4LSL);
            listOfLslUsl.Add(Scara2FittingPins4USL);
            listOfAV.Add(Scara2FittingPins4AV);

            listOfLslUsl.Add(Scara3FittingPins0LSL);
            listOfLslUsl.Add(Scara3FittingPins0USL);
            listOfAV.Add(Scara3FittingPins0AV);

            listOfLslUsl.Add(Scara3FittingPins1LSL);
            listOfLslUsl.Add(Scara3FittingPins1USL);
            listOfAV.Add(Scara3FittingPins1AV);

            listOfLslUsl.Add(Scara3FittingPins2LSL);
            listOfLslUsl.Add(Scara3FittingPins2USL);
            listOfAV.Add(Scara3FittingPins2AV);

            listOfLslUsl.Add(Scara3FittingPins3LSL);
            listOfLslUsl.Add(Scara3FittingPins3USL);
            listOfAV.Add(Scara3FittingPins3AV);

            listOfLslUsl.Add(Scara3FittingPins4LSL);
            listOfLslUsl.Add(Scara3FittingPins4USL);
            listOfAV.Add(Scara3FittingPins4AV);


            txt_rangeValLSL.Add(txt_rangeWeldingDistanceValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeWeldingDistanceValUSL.Text);
            txt_rangeValLSL.Add(txt_rangeWeldingEnergyValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeWeldingEnergyValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeAirFlowFlowVolumeValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeAirFlowFlowVolumeValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara1FittingPins0ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara1FittingPins0ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara1FittingPins1ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara1FittingPins1ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara1FittingPins2ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara1FittingPins2ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara1FittingPins3ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara1FittingPins3ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara1FittingPins4ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara1FittingPins4ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara2FittingPins0ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara2FittingPins0ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara2FittingPins1ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara2FittingPins1ValUSL.Text);


            txt_rangeValLSL.Add(txt_rangeScara2FittingPins2ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara2FittingPins2ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara2FittingPins3ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara2FittingPins3ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara2FittingPins4ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara2FittingPins4ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara3FittingPins0ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara3FittingPins0ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara3FittingPins1ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara3FittingPins1ValUSL.Text);


            txt_rangeValLSL.Add(txt_rangeScara3FittingPins2ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara3FittingPins2ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara3FittingPins3ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara3FittingPins3ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScara3FittingPins4ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScara3FittingPins4ValUSL.Text);


            // Se folosesc pentru metoda ComparareValoriAV
            lbl_status.Add(lbl_statusWeldingDistanceAV);
            lbl_status.Add(lbl_statusWeldingEnergyAV);
            lbl_status.Add(lbl_statusAirFlowFlowVolumeAV);
            lbl_status.Add(lbl_statusScara1FittingPins0AV);
            lbl_status.Add(lbl_statusScara1FittingPins1AV);
            lbl_status.Add(lbl_statusScara1FittingPins2AV);
            lbl_status.Add(lbl_statusScara1FittingPins3AV);
            lbl_status.Add(lbl_statusScara1FittingPins4AV);
            lbl_status.Add(lbl_statusScara2FittingPins0AV);
            lbl_status.Add(lbl_statusScara2FittingPins1AV);
            lbl_status.Add(lbl_statusScara2FittingPins2AV);
            lbl_status.Add(lbl_statusScara2FittingPins3AV);
            lbl_status.Add(lbl_statusScara2FittingPins4AV);
            lbl_status.Add(lbl_statusScara3FittingPins0AV);
            lbl_status.Add(lbl_statusScara3FittingPins1AV);
            lbl_status.Add(lbl_statusScara3FittingPins2AV);
            lbl_status.Add(lbl_statusScara3FittingPins3AV);
            lbl_status.Add(lbl_statusScara3FittingPins4AV);


            lbl_statusBackColor.Add(lbl_statusWeldingDistanceAV);
            lbl_statusBackColor.Add(lbl_statusWeldingEnergyAV);
            lbl_statusBackColor.Add(lbl_statusAirFlowFlowVolumeAV);
            lbl_statusBackColor.Add(lbl_statusScara1FittingPins0AV);
            lbl_statusBackColor.Add(lbl_statusScara1FittingPins1AV);
            lbl_statusBackColor.Add(lbl_statusScara1FittingPins2AV);
            lbl_statusBackColor.Add(lbl_statusScara1FittingPins3AV);
            lbl_statusBackColor.Add(lbl_statusScara1FittingPins4AV);
            lbl_statusBackColor.Add(lbl_statusScara2FittingPins0AV);
            lbl_statusBackColor.Add(lbl_statusScara2FittingPins1AV);
            lbl_statusBackColor.Add(lbl_statusScara2FittingPins2AV);
            lbl_statusBackColor.Add(lbl_statusScara2FittingPins3AV);
            lbl_statusBackColor.Add(lbl_statusScara2FittingPins4AV);
            lbl_statusBackColor.Add(lbl_statusScara3FittingPins0AV);
            lbl_statusBackColor.Add(lbl_statusScara3FittingPins1AV);
            lbl_statusBackColor.Add(lbl_statusScara3FittingPins2AV);
            lbl_statusBackColor.Add(lbl_statusScara3FittingPins3AV);
            lbl_statusBackColor.Add(lbl_statusScara3FittingPins4AV);



            // citirea din fisierul excell a valorilor AV

            for (int i = 0; i < listOfAV.Count; i++)
            {
                ReadExcelFile(rangeDeCititAV[i], i);
            }
            txt_rangeWeldingDistanceValAV.Text = WeldingdistanceAV[0];
            txt_rangeWeldingEnergyValAV.Text = WeldingEnergyAV[0];
            txt_rangeAirFlowFlowVolumeValAV.Text = AirFlowFlowVolumeAV[0];
            txt_rangeScara1FittingPins0ValAV.Text = Scara1FittingPins0AV[0];
            txt_rangeScara1FittingPins1ValAV.Text = Scara1FittingPins1AV[0];
            txt_rangeScara1FittingPins2ValAV.Text = Scara1FittingPins2AV[0];
            txt_rangeScara1FittingPins3ValAV.Text = Scara1FittingPins3AV[0];
            txt_rangeScara1FittingPins4ValAV.Text = Scara1FittingPins4AV[0];
            txt_rangeScara2FittingPins0ValAV.Text = Scara2FittingPins0AV[0];
            txt_rangeScara2FittingPins1ValAV.Text = Scara2FittingPins1AV[0];
            txt_rangeScara2FittingPins2ValAV.Text = Scara2FittingPins2AV[0];
            txt_rangeScara2FittingPins3ValAV.Text = Scara2FittingPins3AV[0];
            txt_rangeScara2FittingPins4ValAV.Text = Scara2FittingPins4AV[0];
            txt_rangeScara3FittingPins0ValAV.Text = Scara3FittingPins0AV[0];
            txt_rangeScara3FittingPins1ValAV.Text = Scara3FittingPins1AV[0];
            txt_rangeScara3FittingPins2ValAV.Text = Scara3FittingPins2AV[0];
            txt_rangeScara3FittingPins3ValAV.Text = Scara3FittingPins3AV[0];
            txt_rangeScara3FittingPins4ValAV.Text = Scara3FittingPins4AV[0];


            ComparareValoriAV();

            // citirea din fisierul excell a valorilor LSL si USL


            for (int i = 0; i < listOfLslUsl.Count; i++)
            {
                ReadExcelFile4(rangeDeCititLslUsl[i],i);
            }
            ComparareValoriWeldingdistanceLSL();
            ComparareValoriWeldingdistanceUSL();
            ComparareValoriWeldingEnergyLSL();
            ComparareValoriWeldingEnergyUSL();
            ComparareValoriAirFlowFlowVolumeLSL();
            ComparareValoriAirFlowFlowVolumeUSL();
            ComparareValoriScara1FittingPins0LSL();
            ComparareValoriScara1FittingPins0USL();
            ComparareValoriScara1FittingPins1LSL();
            ComparareValoriScara1FittingPins1USL();
            ComparareValoriScara1FittingPins2LSL();
            ComparareValoriScara1FittingPins2USL();

            ComparareValoriScara1FittingPins3LSL();
            ComparareValoriScara1FittingPins3USL();

            ComparareValoriScara1FittingPins4LSL();
            ComparareValoriScara1FittingPins4USL();

            ComparareValoriScara2FittingPins0LSL();
            ComparareValoriScara2FittingPins0USL();

            ComparareValoriScara2FittingPins1LSL();
            ComparareValoriScara2FittingPins1USL();

            ComparareValoriScara2FittingPins2LSL();
            ComparareValoriScara2FittingPins2USL();

            ComparareValoriScara2FittingPins3LSL();
            ComparareValoriScara2FittingPins3USL();

            ComparareValoriScara2FittingPins4LSL();
            ComparareValoriScara2FittingPins4USL();
            ComparareValoriScara3FittingPins0LSL();
            ComparareValoriScara3FittingPins0USL();

            ComparareValoriScara3FittingPins1LSL();
            ComparareValoriScara3FittingPins1USL();

            ComparareValoriScara3FittingPins2LSL();
            ComparareValoriScara3FittingPins2USL();

            ComparareValoriScara3FittingPins3LSL();
            ComparareValoriScara3FittingPins3USL();

            ComparareValoriScara3FittingPins4LSL();
            ComparareValoriScara3FittingPins4USL();

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
        public  void ReadExcelFile(string rangeDeCititAV,int i)
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
        public  void ReadExcelFile4(string rangeDeCitit, int i)
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
            

            string WeldingdistanceAVRange = Properties.Settings.Default.txt_rangeWeldingDistance + txt_mediacoloana.Text;
            txt_rangeWeldingDistance.Text = Properties.Settings.Default.txt_rangeWeldingDistance + txt_mediacoloana.Text;
            string WeldingdistanceLSLRange = Properties.Settings.Default.txt_rangeWeldingDistanceLSL + txt_mediacoloana.Text;
            txt_rangeWeldingDistanceLSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceLSL + txt_mediacoloana.Text;
            string WeldingdistanceUSLRange = Properties.Settings.Default.txt_rangeWeldingDistanceUSL + txt_mediacoloana.Text;
            txt_rangeWeldingDistanceUSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceUSL + txt_mediacoloana.Text;

            string WeldingEnergyAVRange = Properties.Settings.Default.txt_rangeWeldingEnergy + txt_mediacoloana.Text;
            txt_rangeWeldingEnergy.Text = Properties.Settings.Default.txt_rangeWeldingEnergy + txt_mediacoloana.Text;
            string WeldinEnergyLSLRange = Properties.Settings.Default.txt_rangeWeldingEnergyLSL + txt_mediacoloana.Text;
            txt_rangeWeldingEnergyLSL.Text = Properties.Settings.Default.txt_rangeWeldingEnergyLSL + txt_mediacoloana.Text;
            string WeldingEnergyUSLRange = Properties.Settings.Default.txt_rangeWeldingEnergyUSL + txt_mediacoloana.Text;
            txt_rangeWeldingEnergyUSL.Text = Properties.Settings.Default.txt_rangeWeldingEnergyUSL + txt_mediacoloana.Text;

            string AirFlowFlowVolumeAVRange = Properties.Settings.Default.txt_rangeAirFlowFlowVolume + txt_mediacoloana.Text;
            txt_rangeAirFlowFlowVolume.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolume + txt_mediacoloana.Text;
            string AirFlowFlowVolumeLSLRange = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeLSL + txt_mediacoloana.Text;
            txt_rangeAirFlowFlowVolumeLSL.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeLSL + txt_mediacoloana.Text;
            string AirFlowFlowVolumeUSLRange = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeUSL + txt_mediacoloana.Text;
            txt_rangeAirFlowFlowVolumeUSL.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeUSL + txt_mediacoloana.Text;


            string Scara1FittingPins0AVRange = Properties.Settings.Default.txt_rangeScara1FittingPins0 + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins0.Text = Properties.Settings.Default.txt_rangeScara1FittingPins0 + txt_mediacoloana.Text;
            string Scara1FittingPins0LSLRange = Properties.Settings.Default.txt_rangeScara1FittingPins0LSL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins0LSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins0LSL + txt_mediacoloana.Text;
            string Scara1FittingPins0USLRange = Properties.Settings.Default.txt_rangeScara1FittingPins0USL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins0USL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins0USL + txt_mediacoloana.Text;

            string Scara1FittingPins1AVRange = Properties.Settings.Default.txt_rangeScara1FittingPins1 + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins1.Text = Properties.Settings.Default.txt_rangeScara1FittingPins1 + txt_mediacoloana.Text;
            string Scara1FittingPins1LSLRange = Properties.Settings.Default.txt_rangeScara1FittingPins1LSL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins1LSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins1LSL + txt_mediacoloana.Text;
            string Scara1FittingPins1USLRange = Properties.Settings.Default.txt_rangeScara1FittingPins1USL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins1USL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins1USL + txt_mediacoloana.Text;

            string Scara1FittingPins2AVRange = Properties.Settings.Default.txt_rangeScara1FittingPins2 + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins2.Text = Properties.Settings.Default.txt_rangeScara1FittingPins2 + txt_mediacoloana.Text;
            string Scara1FittingPins2LSLRange = Properties.Settings.Default.txt_rangeScara1FittingPins2LSL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins2LSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins2LSL + txt_mediacoloana.Text;
            string Scara1FittingPins2USLRange = Properties.Settings.Default.txt_rangeScara1FittingPins2USL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins2USL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins2USL + txt_mediacoloana.Text;

            string Scara1FittingPins3AVRange = Properties.Settings.Default.txt_rangeScara1FittingPins3 + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins3.Text = Properties.Settings.Default.txt_rangeScara1FittingPins3 + txt_mediacoloana.Text;
            string Scara1FittingPins3LSLRange = Properties.Settings.Default.txt_rangeScara1FittingPins3LSL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins3LSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins3LSL + txt_mediacoloana.Text;
            string Scara1FittingPins3USLRange = Properties.Settings.Default.txt_rangeScara1FittingPins3USL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins3USL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins3USL + txt_mediacoloana.Text;

            string Scara1FittingPins4AVRange = Properties.Settings.Default.txt_rangeScara1FittingPins4 + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins4.Text = Properties.Settings.Default.txt_rangeScara1FittingPins4 + txt_mediacoloana.Text;
            string Scara1FittingPins4LSLRange = Properties.Settings.Default.txt_rangeScara1FittingPins4LSL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins4LSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins4LSL + txt_mediacoloana.Text;
            string Scara1FittingPins4USLRange = Properties.Settings.Default.txt_rangeScara1FittingPins4USL + txt_mediacoloana.Text;
            txt_rangeScara1FittingPins4USL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins4USL + txt_mediacoloana.Text;

            string Scara2FittingPins0AVRange = Properties.Settings.Default.txt_rangeScara2FittingPins0 + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins0.Text = Properties.Settings.Default.txt_rangeScara2FittingPins0 + txt_mediacoloana.Text;
            string Scara2FittingPins0LSLRange = Properties.Settings.Default.txt_rangeScara2FittingPins0LSL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins0LSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins0LSL + txt_mediacoloana.Text;
            string Scara2FittingPins0USLRange = Properties.Settings.Default.txt_rangeScara2FittingPins0USL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins0USL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins0USL + txt_mediacoloana.Text;

            string Scara2FittingPins1AVRange = Properties.Settings.Default.txt_rangeScara2FittingPins1 + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins1.Text = Properties.Settings.Default.txt_rangeScara2FittingPins1 + txt_mediacoloana.Text;
            string Scara2FittingPins1LSLRange = Properties.Settings.Default.txt_rangeScara2FittingPins1LSL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins1LSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins1LSL + txt_mediacoloana.Text;
            string Scara2FittingPins1USLRange = Properties.Settings.Default.txt_rangeScara2FittingPins1USL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins1USL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins1USL + txt_mediacoloana.Text;

            string Scara2FittingPins2AVRange = Properties.Settings.Default.txt_rangeScara2FittingPins2 + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins2.Text = Properties.Settings.Default.txt_rangeScara2FittingPins2 + txt_mediacoloana.Text;
            string Scara2FittingPins2LSLRange = Properties.Settings.Default.txt_rangeScara2FittingPins2LSL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins2LSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins2LSL + txt_mediacoloana.Text;
            string Scara2FittingPins2USLRange = Properties.Settings.Default.txt_rangeScara2FittingPins2USL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins2USL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins2USL + txt_mediacoloana.Text;

            string Scara2FittingPins3AVRange = Properties.Settings.Default.txt_rangeScara2FittingPins3 + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins3.Text = Properties.Settings.Default.txt_rangeScara2FittingPins3 + txt_mediacoloana.Text;
            string Scara2FittingPins3LSLRange = Properties.Settings.Default.txt_rangeScara2FittingPins3LSL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins3LSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins3LSL + txt_mediacoloana.Text;
            string Scara2FittingPins3USLRange = Properties.Settings.Default.txt_rangeScara2FittingPins3USL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins3USL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins3USL + txt_mediacoloana.Text;

            string Scara2FittingPins4AVRange = Properties.Settings.Default.txt_rangeScara2FittingPins4 + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins4.Text = Properties.Settings.Default.txt_rangeScara2FittingPins4 + txt_mediacoloana.Text;
            string Scara2FittingPins4LSLRange = Properties.Settings.Default.txt_rangeScara2FittingPins4LSL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins4LSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins4LSL + txt_mediacoloana.Text;
            string Scara2FittingPins4USLRange = Properties.Settings.Default.txt_rangeScara2FittingPins4USL + txt_mediacoloana.Text;
            txt_rangeScara2FittingPins4USL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins4USL + txt_mediacoloana.Text;

            string Scara3FittingPins0AVRange = Properties.Settings.Default.txt_rangeScara3FittingPins0 + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins0.Text = Properties.Settings.Default.txt_rangeScara3FittingPins0 + txt_mediacoloana.Text;
            string Scara3FittingPins0LSLRange = Properties.Settings.Default.txt_rangeScara3FittingPins0LSL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins0LSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins0LSL + txt_mediacoloana.Text;
            string Scara3FittingPins0USLRange = Properties.Settings.Default.txt_rangeScara3FittingPins0USL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins0USL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins0USL + txt_mediacoloana.Text;

            string Scara3FittingPins1AVRange = Properties.Settings.Default.txt_rangeScara3FittingPins1 + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins1.Text = Properties.Settings.Default.txt_rangeScara3FittingPins1 + txt_mediacoloana.Text;
            string Scara3FittingPins1LSLRange = Properties.Settings.Default.txt_rangeScara3FittingPins1LSL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins1LSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins1LSL + txt_mediacoloana.Text;
            string Scara3FittingPins1USLRange = Properties.Settings.Default.txt_rangeScara3FittingPins1USL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins1USL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins1USL + txt_mediacoloana.Text;

            string Scara3FittingPins2AVRange = Properties.Settings.Default.txt_rangeScara3FittingPins2 + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins2.Text = Properties.Settings.Default.txt_rangeScara3FittingPins2 + txt_mediacoloana.Text;
            string Scara3FittingPins2LSLRange = Properties.Settings.Default.txt_rangeScara3FittingPins2LSL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins2LSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins2LSL + txt_mediacoloana.Text;
            string Scara3FittingPins2USLRange = Properties.Settings.Default.txt_rangeScara3FittingPins2USL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins2USL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins2USL + txt_mediacoloana.Text;

            string Scara3FittingPins3AVRange = Properties.Settings.Default.txt_rangeScara3FittingPins3 + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins3.Text = Properties.Settings.Default.txt_rangeScara3FittingPins3 + txt_mediacoloana.Text;
            string Scara3FittingPins3LSLRange = Properties.Settings.Default.txt_rangeScara3FittingPins3LSL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins3LSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins3LSL + txt_mediacoloana.Text;
            string Scara3FittingPins3USLRange = Properties.Settings.Default.txt_rangeScara3FittingPins3USL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins3USL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins3USL + txt_mediacoloana.Text;

            string Scara3FittingPins4AVRange = Properties.Settings.Default.txt_rangeScara3FittingPins4 + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins4.Text = Properties.Settings.Default.txt_rangeScara3FittingPins4 + txt_mediacoloana.Text;
            string Scara3FittingPins4LSLRange = Properties.Settings.Default.txt_rangeScara3FittingPins4LSL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins4LSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins4LSL + txt_mediacoloana.Text;
            string Scara3FittingPins4USLRange = Properties.Settings.Default.txt_rangeScara3FittingPins4USL + txt_mediacoloana.Text;
            txt_rangeScara3FittingPins4USL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins4USL + txt_mediacoloana.Text;



            /// welding distance av range
            rangeDeCititAV.Add(WeldingdistanceAVRange);
            rangeDeCititLslUsl.Add(WeldingdistanceLSLRange);
            rangeDeCititLslUsl.Add(WeldingdistanceUSLRange);
            // welding energy av range
            rangeDeCititAV.Add(WeldingEnergyAVRange);
            rangeDeCititLslUsl.Add(WeldinEnergyLSLRange);
            rangeDeCititLslUsl.Add(WeldingEnergyUSLRange);
            // AirFlowFlowVolume av range
            rangeDeCititAV.Add(AirFlowFlowVolumeAVRange);
            rangeDeCititLslUsl.Add(AirFlowFlowVolumeLSLRange);
            rangeDeCititLslUsl.Add(AirFlowFlowVolumeUSLRange);

            // ST21 Scara1FittingPins0 av range
            rangeDeCititAV.Add(Scara1FittingPins0AVRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins0LSLRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins0USLRange);

            // ST21 Scara1FittingPins1 av range
            rangeDeCititAV.Add(Scara1FittingPins1AVRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins1LSLRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins1USLRange);

            // ST21 Scara1FittingPins2 av range
            rangeDeCititAV.Add(Scara1FittingPins2AVRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins2LSLRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins2USLRange);

            // ST21 Scara1FittingPins3 av range
            rangeDeCititAV.Add(Scara1FittingPins3AVRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins3LSLRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins3USLRange);

            // ST21 Scara1FittingPins4 av range
            rangeDeCititAV.Add(Scara1FittingPins4AVRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins4LSLRange);
            rangeDeCititLslUsl.Add(Scara1FittingPins4USLRange);

            // ST21 Scara2FittingPins0 av range
            rangeDeCititAV.Add(Scara2FittingPins0AVRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins0LSLRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins0USLRange);

            // ST21 Scara2FittingPins1 av range
            rangeDeCititAV.Add(Scara2FittingPins1AVRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins1LSLRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins1USLRange);

            // ST21 Scara2FittingPins2 av range
            rangeDeCititAV.Add(Scara2FittingPins2AVRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins2LSLRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins2USLRange);

            // ST21 Scara2FittingPins3 av range
            rangeDeCititAV.Add(Scara2FittingPins3AVRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins3LSLRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins3USLRange);

            // ST21 Scara2FittingPins4 av range
            rangeDeCititAV.Add(Scara2FittingPins4AVRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins4LSLRange);
            rangeDeCititLslUsl.Add(Scara2FittingPins4USLRange);

            // ST21 Scara3FittingPins0 av range
            rangeDeCititAV.Add(Scara3FittingPins0AVRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins0LSLRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins0USLRange);

            // ST21 Scara3FittingPins1 av range
            rangeDeCititAV.Add(Scara3FittingPins1AVRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins1LSLRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins1USLRange);

            // ST21 Scara3FittingPins2 av range
            rangeDeCititAV.Add(Scara3FittingPins2AVRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins2LSLRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins2USLRange);

            // ST21 Scara3FittingPins3 av range
            rangeDeCititAV.Add(Scara3FittingPins3AVRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins3LSLRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins3USLRange);

            // ST21 Scara3FittingPins4 av range
            rangeDeCititAV.Add(Scara3FittingPins4AVRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins4LSLRange);
            rangeDeCititLslUsl.Add(Scara3FittingPins4USLRange);


        }
        public void SetareValoriCampuriFisiere()
        {
            txt_logfile.Text = Properties.Settings.Default.txt_logfile;
            txt_rangeWeldingDistanceValLSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceValLSL.ToString();
            txt_rangeWeldingDistanceValUSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceValUSL.ToString();

            txt_rangeWeldingEnergyValLSL.Text = Properties.Settings.Default.txt_rangeWeldingEnergyValLSL.ToString();
            txt_rangeWeldingEnergyValUSL.Text = Properties.Settings.Default.txt_rangeWeldingEnergyValUSL.ToString();

            txt_rangeAirFlowFlowVolumeValLSL.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeValLSL.ToString();
            txt_rangeAirFlowFlowVolumeValUSL.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeValUSL.ToString();

            txt_rangeScara1FittingPins0ValLSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins0ValLSL.ToString();
            txt_rangeScara1FittingPins0ValUSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins0ValUSL.ToString();

            txt_rangeScara1FittingPins1ValLSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins1ValLSL.ToString();
            txt_rangeScara1FittingPins1ValUSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins1ValUSL.ToString();

            txt_rangeScara1FittingPins2ValLSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins2ValLSL.ToString();
            txt_rangeScara1FittingPins2ValUSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins2ValUSL.ToString();

            txt_rangeScara1FittingPins3ValLSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins3ValLSL.ToString();
            txt_rangeScara1FittingPins3ValUSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins3ValUSL.ToString();

            txt_rangeScara1FittingPins4ValLSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins4ValLSL.ToString();
            txt_rangeScara1FittingPins4ValUSL.Text = Properties.Settings.Default.txt_rangeScara1FittingPins4ValUSL.ToString();

            txt_rangeScara2FittingPins0ValLSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins0ValLSL.ToString();
            txt_rangeScara2FittingPins0ValUSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins0ValUSL.ToString();

            txt_rangeScara2FittingPins1ValLSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins1ValLSL.ToString();
            txt_rangeScara2FittingPins1ValUSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins1ValUSL.ToString();

            txt_rangeScara2FittingPins2ValLSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins2ValLSL.ToString();
            txt_rangeScara2FittingPins2ValUSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins2ValUSL.ToString();

            txt_rangeScara2FittingPins3ValLSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins3ValLSL.ToString();
            txt_rangeScara2FittingPins3ValUSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins3ValUSL.ToString();

            txt_rangeScara2FittingPins4ValLSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins4ValLSL.ToString();
            txt_rangeScara2FittingPins4ValUSL.Text = Properties.Settings.Default.txt_rangeScara2FittingPins4ValUSL.ToString();

            txt_rangeScara3FittingPins0ValLSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins0ValLSL.ToString();
            txt_rangeScara3FittingPins0ValUSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins0ValUSL.ToString();

            txt_rangeScara3FittingPins1ValLSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins1ValLSL.ToString();
            txt_rangeScara3FittingPins1ValUSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins1ValUSL.ToString();

            txt_rangeScara3FittingPins2ValLSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins2ValLSL.ToString();
            txt_rangeScara3FittingPins2ValUSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins2ValUSL.ToString();

            txt_rangeScara3FittingPins3ValLSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins3ValLSL.ToString();
            txt_rangeScara3FittingPins3ValUSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins3ValUSL.ToString();

            txt_rangeScara3FittingPins4ValLSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins4ValLSL.ToString();
            txt_rangeScara3FittingPins4ValUSL.Text = Properties.Settings.Default.txt_rangeScara3FittingPins4ValUSL.ToString();

        }

        // pentru a compara valorile AV si a afisa OK sau NOK

        public void ComparareValoriAV()
        {
            for (int i = 0; i < listOfAV.Count; i++)
            {
                for (int j = 0; j < listOfAV[i].Count; j++)
                {
                    if (Convert.ToDouble(listOfAV[i][j]) >= Convert.ToDouble(txt_rangeValLSL[i].ToString()) && Convert.ToDouble(listOfAV[i][j]) <= Convert.ToDouble(txt_rangeValUSL[i].ToString()))
                    {
                        lbl_status[i].Text = "OK";
                        lbl_statusBackColor[i].BackColor = Color.GreenYellow;
                    }

                    else
                    {
                        lbl_status[i].Text = "NOK";
                        lbl_statusBackColor[i].BackColor = Color.Red;
                        break;
                    }
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

       

        public void ComparareValoriWeldingEnergyLSL()
        {
            for (int i = 0; i < WeldingEnergyLSL.Count; i++)
            {
                if (WeldingEnergyLSL[i].Equals(txt_rangeWeldingEnergyValLSL.Text.ToString()))
                {
                    lbl_statusWeldingEnergyLSL.Text = "OK";
                    lbl_statusWeldingEnergyLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusWeldingEnergyLSL.Text = "NOK";
                    lbl_statusWeldingEnergyLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriWeldingEnergyUSL()
        {
            for (int i = 0; i < WeldingEnergyUSL.Count; i++)
            {
                if (WeldingEnergyUSL[i].Equals(txt_rangeWeldingEnergyValUSL.Text.ToString()))
                {
                    lbl_statusWeldingEnergyUSL.Text = "OK";
                    lbl_statusWeldingEnergyUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusWeldingEnergyUSL.Text = "NOK";
                    lbl_statusWeldingEnergyUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriAirFlowFlowVolumeLSL()
        {
            for (int i = 0; i < AirFlowFlowVolumeLSL.Count; i++)
            {
                if (AirFlowFlowVolumeLSL[i].Equals(txt_rangeAirFlowFlowVolumeValLSL.Text.ToString()))
                {
                    lbl_statusAirFlowFlowVolumeLSL.Text = "OK";
                    lbl_statusAirFlowFlowVolumeLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusAirFlowFlowVolumeLSL.Text = "NOK";
                    lbl_statusAirFlowFlowVolumeLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriAirFlowFlowVolumeUSL()
        {
            for (int i = 0; i < AirFlowFlowVolumeUSL.Count; i++)
            {
                if (AirFlowFlowVolumeUSL[i].Equals(txt_rangeAirFlowFlowVolumeValUSL.Text.ToString()))
                {
                    lbl_statusAirFlowFlowVolumeUSL.Text = "OK";
                    lbl_statusAirFlowFlowVolumeUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusAirFlowFlowVolumeUSL.Text = "NOK";
                    lbl_statusAirFlowFlowVolumeUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins0LSL()
        {
            for (int i = 0; i < Scara1FittingPins0LSL.Count; i++)
            {
                if (Scara1FittingPins0LSL[i].Equals(txt_rangeScara1FittingPins0ValLSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins0LSL.Text = "OK";
                    lbl_statusScara1FittingPins0LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins0LSL.Text = "NOK";
                    lbl_statusScara1FittingPins0LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins0USL()
        {
            for (int i = 0; i < Scara1FittingPins0USL.Count; i++)
            {
                if (Scara1FittingPins0USL[i].Equals(txt_rangeScara1FittingPins0ValUSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins0USL.Text = "OK";
                    lbl_statusScara1FittingPins0USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins0USL.Text = "NOK";
                    lbl_statusScara1FittingPins0USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins1LSL()
        {
            for (int i = 0; i < Scara1FittingPins1LSL.Count; i++)
            {
                if (Scara1FittingPins1LSL[i].Equals(txt_rangeScara1FittingPins1ValLSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins1LSL.Text = "OK";
                    lbl_statusScara1FittingPins1LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins1LSL.Text = "NOK";
                    lbl_statusScara1FittingPins1LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins1USL()
        {
            for (int i = 0; i < Scara1FittingPins1USL.Count; i++)
            {
                if (Scara1FittingPins1USL[i].Equals(txt_rangeScara1FittingPins1ValUSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins1USL.Text = "OK";
                    lbl_statusScara1FittingPins1USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins1USL.Text = "NOK";
                    lbl_statusScara1FittingPins1USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins2LSL()
        {
            for (int i = 0; i < Scara1FittingPins2LSL.Count; i++)
            {
                if (Scara1FittingPins2LSL[i].Equals(txt_rangeScara1FittingPins2ValLSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins2LSL.Text = "OK";
                    lbl_statusScara1FittingPins2LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins2LSL.Text = "NOK";
                    lbl_statusScara1FittingPins2LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins2USL()
        {
            for (int i = 0; i < Scara1FittingPins2USL.Count; i++)
            {
                if (Scara1FittingPins2USL[i].Equals(txt_rangeScara1FittingPins2ValUSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins2USL.Text = "OK";
                    lbl_statusScara1FittingPins2USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins2USL.Text = "NOK";
                    lbl_statusScara1FittingPins2USL.BackColor = Color.Red;
                    break;
                }
            }

        }


        public void ComparareValoriScara1FittingPins3LSL()
        {
            for (int i = 0; i < Scara1FittingPins3LSL.Count; i++)
            {
                if (Scara1FittingPins3LSL[i].Equals(txt_rangeScara1FittingPins3ValLSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins3LSL.Text = "OK";
                    lbl_statusScara1FittingPins3LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins3LSL.Text = "NOK";
                    lbl_statusScara1FittingPins3LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins3USL()
        {
            for (int i = 0; i < Scara1FittingPins3USL.Count; i++)
            {
                if (Scara1FittingPins3USL[i].Equals(txt_rangeScara1FittingPins3ValUSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins3USL.Text = "OK";
                    lbl_statusScara1FittingPins3USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins3USL.Text = "NOK";
                    lbl_statusScara1FittingPins3USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins4LSL()
        {
            for (int i = 0; i < Scara1FittingPins4LSL.Count; i++)
            {
                if (Scara1FittingPins4LSL[i].Equals(txt_rangeScara1FittingPins4ValLSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins4LSL.Text = "OK";
                    lbl_statusScara1FittingPins4LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins4LSL.Text = "NOK";
                    lbl_statusScara1FittingPins4LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara1FittingPins4USL()
        {
            for (int i = 0; i < Scara1FittingPins4USL.Count; i++)
            {
                if (Scara1FittingPins4USL[i].Equals(txt_rangeScara1FittingPins4ValUSL.Text.ToString()))
                {
                    lbl_statusScara1FittingPins4USL.Text = "OK";
                    lbl_statusScara1FittingPins4USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara1FittingPins4USL.Text = "NOK";
                    lbl_statusScara1FittingPins4USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins0LSL()
        {
            for (int i = 0; i < Scara2FittingPins0LSL.Count; i++)
            {
                if (Scara2FittingPins0LSL[i].Equals(txt_rangeScara2FittingPins0ValLSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins0LSL.Text = "OK";
                    lbl_statusScara2FittingPins0LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins0LSL.Text = "NOK";
                    lbl_statusScara2FittingPins0LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins0USL()
        {
            for (int i = 0; i < Scara2FittingPins0USL.Count; i++)
            {
                if (Scara2FittingPins0USL[i].Equals(txt_rangeScara2FittingPins0ValUSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins0USL.Text = "OK";
                    lbl_statusScara2FittingPins0USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins0USL.Text = "NOK";
                    lbl_statusScara2FittingPins0USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriScara2FittingPins1LSL()
        {
            for (int i = 0; i < Scara2FittingPins1LSL.Count; i++)
            {
                if (Scara2FittingPins1LSL[i].Equals(txt_rangeScara2FittingPins1ValLSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins1LSL.Text = "OK";
                    lbl_statusScara2FittingPins1LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins1LSL.Text = "NOK";
                    lbl_statusScara2FittingPins1LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins1USL()
        {
            for (int i = 0; i < Scara2FittingPins1USL.Count; i++)
            {
                if (Scara2FittingPins1USL[i].Equals(txt_rangeScara2FittingPins1ValUSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins1USL.Text = "OK";
                    lbl_statusScara2FittingPins1USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins1USL.Text = "NOK";
                    lbl_statusScara2FittingPins1USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriScara2FittingPins2LSL()
        {
            for (int i = 0; i < Scara2FittingPins2LSL.Count; i++)
            {
                if (Scara2FittingPins2LSL[i].Equals(txt_rangeScara2FittingPins2ValLSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins2LSL.Text = "OK";
                    lbl_statusScara2FittingPins2LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins2LSL.Text = "NOK";
                    lbl_statusScara2FittingPins2LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins2USL()
        {
            for (int i = 0; i < Scara2FittingPins2USL.Count; i++)
            {
                if (Scara2FittingPins2USL[i].Equals(txt_rangeScara2FittingPins2ValUSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins2USL.Text = "OK";
                    lbl_statusScara2FittingPins2USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins2USL.Text = "NOK";
                    lbl_statusScara2FittingPins2USL.BackColor = Color.Red;
                    break;
                }
            }

        }


        public void ComparareValoriScara2FittingPins3LSL()
        {
            for (int i = 0; i < Scara2FittingPins3LSL.Count; i++)
            {
                if (Scara2FittingPins3LSL[i].Equals(txt_rangeScara2FittingPins3ValLSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins3LSL.Text = "OK";
                    lbl_statusScara2FittingPins3LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins3LSL.Text = "NOK";
                    lbl_statusScara2FittingPins3LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins3USL()
        {
            for (int i = 0; i < Scara2FittingPins3USL.Count; i++)
            {
                if (Scara2FittingPins3USL[i].Equals(txt_rangeScara2FittingPins3ValUSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins3USL.Text = "OK";
                    lbl_statusScara2FittingPins3USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins3USL.Text = "NOK";
                    lbl_statusScara2FittingPins3USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins4LSL()
        {
            for (int i = 0; i < Scara2FittingPins4LSL.Count; i++)
            {
                if (Scara2FittingPins4LSL[i].Equals(txt_rangeScara2FittingPins4ValLSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins4LSL.Text = "OK";
                    lbl_statusScara2FittingPins4LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins4LSL.Text = "NOK";
                    lbl_statusScara2FittingPins4LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara2FittingPins4USL()
        {
            for (int i = 0; i < Scara2FittingPins4USL.Count; i++)
            {
                if (Scara2FittingPins4USL[i].Equals(txt_rangeScara2FittingPins4ValUSL.Text.ToString()))
                {
                    lbl_statusScara2FittingPins4USL.Text = "OK";
                    lbl_statusScara2FittingPins4USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara2FittingPins4USL.Text = "NOK";
                    lbl_statusScara2FittingPins4USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins0LSL()
        {
            for (int i = 0; i < Scara3FittingPins0LSL.Count; i++)
            {
                if (Scara3FittingPins0LSL[i].Equals(txt_rangeScara3FittingPins0ValLSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins0LSL.Text = "OK";
                    lbl_statusScara3FittingPins0LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins0LSL.Text = "NOK";
                    lbl_statusScara3FittingPins0LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins0USL()
        {
            for (int i = 0; i < Scara3FittingPins0USL.Count; i++)
            {
                if (Scara3FittingPins0USL[i].Equals(txt_rangeScara3FittingPins0ValUSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins0USL.Text = "OK";
                    lbl_statusScara3FittingPins0USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins0USL.Text = "NOK";
                    lbl_statusScara3FittingPins0USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriScara3FittingPins1LSL()
        {
            for (int i = 0; i < Scara3FittingPins1LSL.Count; i++)
            {
                if (Scara3FittingPins1LSL[i].Equals(txt_rangeScara3FittingPins1ValLSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins1LSL.Text = "OK";
                    lbl_statusScara3FittingPins1LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins1LSL.Text = "NOK";
                    lbl_statusScara3FittingPins1LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins1USL()
        {
            for (int i = 0; i < Scara3FittingPins1USL.Count; i++)
            {
                if (Scara3FittingPins1USL[i].Equals(txt_rangeScara3FittingPins1ValUSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins1USL.Text = "OK";
                    lbl_statusScara3FittingPins1USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins1USL.Text = "NOK";
                    lbl_statusScara3FittingPins1USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriScara3FittingPins2LSL()
        {
            for (int i = 0; i < Scara3FittingPins2LSL.Count; i++)
            {
                if (Scara3FittingPins2LSL[i].Equals(txt_rangeScara3FittingPins2ValLSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins2LSL.Text = "OK";
                    lbl_statusScara3FittingPins2LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins2LSL.Text = "NOK";
                    lbl_statusScara3FittingPins2LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins2USL()
        {
            for (int i = 0; i < Scara3FittingPins2USL.Count; i++)
            {
                if (Scara3FittingPins2USL[i].Equals(txt_rangeScara3FittingPins2ValUSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins2USL.Text = "OK";
                    lbl_statusScara3FittingPins2USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins2USL.Text = "NOK";
                    lbl_statusScara3FittingPins2USL.BackColor = Color.Red;
                    break;
                }
            }

        }


        public void ComparareValoriScara3FittingPins3LSL()
        {
            for (int i = 0; i < Scara3FittingPins3LSL.Count; i++)
            {
                if (Scara3FittingPins3LSL[i].Equals(txt_rangeScara3FittingPins3ValLSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins3LSL.Text = "OK";
                    lbl_statusScara3FittingPins3LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins3LSL.Text = "NOK";
                    lbl_statusScara3FittingPins3LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins3USL()
        {
            for (int i = 0; i < Scara3FittingPins3USL.Count; i++)
            {
                if (Scara3FittingPins3USL[i].Equals(txt_rangeScara3FittingPins3ValUSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins3USL.Text = "OK";
                    lbl_statusScara3FittingPins3USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins3USL.Text = "NOK";
                    lbl_statusScara3FittingPins3USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins4LSL()
        {
            for (int i = 0; i < Scara3FittingPins4LSL.Count; i++)
            {
                if (Scara3FittingPins4LSL[i].Equals(txt_rangeScara3FittingPins4ValLSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins4LSL.Text = "OK";
                    lbl_statusScara3FittingPins4LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins4LSL.Text = "NOK";
                    lbl_statusScara3FittingPins4LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScara3FittingPins4USL()
        {
            for (int i = 0; i < Scara3FittingPins4USL.Count; i++)
            {
                if (Scara3FittingPins4USL[i].Equals(txt_rangeScara3FittingPins4ValUSL.Text.ToString()))
                {
                    lbl_statusScara3FittingPins4USL.Text = "OK";
                    lbl_statusScara3FittingPins4USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScara3FittingPins4USL.Text = "NOK";
                    lbl_statusScara3FittingPins4USL.BackColor = Color.Red;
                    break;
                }
            }

        }



    }
    // nu e introdus
    /*
    St010.Welding.Frequency.AV
    St010.Welding.Weldingforce
    St010.Airflowtest.Time.AV

     */

}


