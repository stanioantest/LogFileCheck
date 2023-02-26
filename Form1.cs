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

            txt_rangeValLSL.Add(txt_rangeWeldingDistanceValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeWeldingDistanceValUSL.Text);
            txt_rangeValLSL.Add(txt_rangeWeldingEnergyValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeWeldingEnergyValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeAirFlowFlowVolumeValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeAirFlowFlowVolumeValUSL.Text);


            // Se folosesc pentru metoda ComparareValoriAV
            lbl_status.Add(lbl_statusWeldingDistanceAV);
            lbl_status.Add(lbl_statusWeldingEnergyAV);
            lbl_status.Add(lbl_statusAirFlowFlowVolumeAV);


            lbl_statusBackColor.Add(lbl_statusWeldingDistanceAV);
            lbl_statusBackColor.Add(lbl_statusWeldingEnergyAV);
            lbl_statusBackColor.Add(lbl_statusAirFlowFlowVolumeAV);



            // citirea din fisierul excell a valorilor AV

            for (int i = 0; i < listOfAV.Count; i++)
            {
                ReadExcelFile(rangeDeCititAV[i], i);
            }
            txt_rangeWeldingDistanceValAV.Text = WeldingdistanceAV[0];
            txt_rangeWeldingEnergyValAV.Text = WeldingEnergyAV[0];

            txt_rangeAirFlowFlowVolumeValAV.Text = AirFlowFlowVolumeAV[0];

            ComparareValoriAV();

          //  ComparareValoriWeldingdistanceAV();
          // ComparareValoriWeldingEnergyAV();

            // citirea din fisierul excell a valorilor LSL si USL
            // 



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

            string WeldingdistanceAVRange = Properties.Settings.Default.txt_rangeWeldingDistance + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingDistance.Text = Properties.Settings.Default.txt_rangeWeldingDistance + AflareRanduriDinExcell().ToString();
            string WeldingdistanceLSLRange = Properties.Settings.Default.txt_rangeWeldingDistanceLSL + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingDistanceLSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceLSL + AflareRanduriDinExcell().ToString();
            string WeldingdistanceUSLRange = Properties.Settings.Default.txt_rangeWeldingDistanceUSL + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingDistanceUSL.Text = Properties.Settings.Default.txt_rangeWeldingDistanceUSL + AflareRanduriDinExcell().ToString();

            string WeldingEnergyAVRange = Properties.Settings.Default.txt_rangeWeldingEnergy + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingEnergy.Text = Properties.Settings.Default.txt_rangeWeldingEnergy + AflareRanduriDinExcell().ToString();
            string WeldinEnergyLSLRange = Properties.Settings.Default.txt_rangeWeldingEnergyLSL + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingEnergyLSL.Text = Properties.Settings.Default.txt_rangeWeldingEnergyLSL + AflareRanduriDinExcell().ToString();
            string WeldingEnergyUSLRange = Properties.Settings.Default.txt_rangeWeldingEnergyUSL + AflareRanduriDinExcell().ToString();
            txt_rangeWeldingEnergyUSL.Text = Properties.Settings.Default.txt_rangeWeldingEnergyUSL + AflareRanduriDinExcell().ToString();

            string AirFlowFlowVolumeAVRange = Properties.Settings.Default.txt_rangeAirFlowFlowVolume + AflareRanduriDinExcell().ToString();
            txt_rangeAirFlowFlowVolume.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolume + AflareRanduriDinExcell().ToString();
            string AirFlowFlowVolumeLSLRange = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeLSL + AflareRanduriDinExcell().ToString();
            txt_rangeAirFlowFlowVolumeLSL.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeLSL + AflareRanduriDinExcell().ToString();
            string AirFlowFlowVolumeUSLRange = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeUSL + AflareRanduriDinExcell().ToString();
            txt_rangeAirFlowFlowVolumeUSL.Text = Properties.Settings.Default.txt_rangeAirFlowFlowVolumeUSL + AflareRanduriDinExcell().ToString();

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


    }
    // nu e introdus
    /*
    St010.Welding.Frequency.AV
    St010.Welding.Weldingforce
    St010.Airflowtest.Time.AV

     */

}


