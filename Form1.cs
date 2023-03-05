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

        List<string> ProfileScanPoint0AV = new List<string>();
        List<string> ProfileScanPoint0LSL = new List<string>();
        List<string> ProfileScanPoint0USL = new List<string>();

        List<string> ProfileScanPoint1AV = new List<string>();
        List<string> ProfileScanPoint1LSL = new List<string>();
        List<string> ProfileScanPoint1USL = new List<string>();

        List<string> ProfileScanPoint2AV = new List<string>();
        List<string> ProfileScanPoint2LSL = new List<string>();
        List<string> ProfileScanPoint2USL = new List<string>();

        List<string> ProfileScanPoint3AV = new List<string>();
        List<string> ProfileScanPoint3LSL = new List<string>();
        List<string> ProfileScanPoint3USL = new List<string>();

        List<string> ProfileScanPoint4AV = new List<string>();
        List<string> ProfileScanPoint4LSL = new List<string>();
        List<string> ProfileScanPoint4USL = new List<string>();

        List<string> ProfileScanPoint5AV = new List<string>();
        List<string> ProfileScanPoint5LSL = new List<string>();
        List<string> ProfileScanPoint5USL = new List<string>();

        List<string> JoiningForceAV = new List<string>();
        List<string> JoiningForceLSL = new List<string>();
        List<string> JoiningForceUSL = new List<string>();

        List<string> LaserWeldingSettingDistanceAV = new List<string>();
        List<string> LaserWeldingSettingDistanceLSL = new List<string>();
        List<string> LaserWeldingSettingDistanceUSL = new List<string>();

        List<string> LaserWeldingStartPositionAV = new List<string>();
        List<string> LaserWeldingStartPositionLSL = new List<string>();
        List<string> LaserWeldingStartPositionUSL = new List<string>();

        List<string> LaserWeldingShutdownDistanceAV = new List<string>();
        List<string> LaserWeldingShutdownDistanceLSL = new List<string>();
        List<string> LaserWeldingShutdownDistanceUSL = new List<string>();

        List<string> LaserWeldingWeldingtimeAV = new List<string>();
        List<string> LaserWeldingWeldingtimeLSL = new List<string>();
        List<string> LaserWeldingWeldingtimeUSL = new List<string>();

        List<string> ForceAV = new List<string>();
        List<string> ForceLSL = new List<string>();
        List<string> ForceUSL = new List<string>();

        List<string> LeaktestStartPressureAV = new List<string>();
        List<string> LeaktestStartPressureLSL = new List<string>();
        List<string> LeaktestStartPressureUSL = new List<string>();

        List<string> LeaktestLeakageAV = new List<string>();
        List<string> LeaktestLeakageLSL = new List<string>();
        List<string> LeaktestLeakageUSL = new List<string>();

        List<string> EOLM01AV = new List<string>();
        List<string> EOLM01LSL = new List<string>();
        List<string> EOLM01USL = new List<string>();

        List<string> EOLM02AV = new List<string>();
        List<string> EOLM02LSL = new List<string>();
        List<string> EOLM02USL = new List<string>();

        List<string> EOLM03AV = new List<string>();
        List<string> EOLM03LSL = new List<string>();
        List<string> EOLM03USL = new List<string>();

        List<string> EOLM04AV = new List<string>();
        List<string> EOLM04LSL = new List<string>();
        List<string> EOLM04USL = new List<string>();

        List<string> EOLM05AV = new List<string>();
        List<string> EOLM05LSL = new List<string>();
        List<string> EOLM05USL = new List<string>();

        List<string> EOLM06AV = new List<string>();
        List<string> EOLM06LSL = new List<string>();
        List<string> EOLM06USL = new List<string>();

        List<string> EOLM07AV = new List<string>();
        List<string> EOLM07LSL = new List<string>();
        List<string> EOLM07USL = new List<string>();

        List<string> EOLM08AV = new List<string>();
        List<string> EOLM08LSL = new List<string>();
        List<string> EOLM08USL = new List<string>();

        List<string> EOLM09AV = new List<string>();
        List<string> EOLM09LSL = new List<string>();
        List<string> EOLM09USL = new List<string>();

        List<string> EOLM10AV = new List<string>();
        List<string> EOLM10LSL = new List<string>();
        List<string> EOLM10USL = new List<string>();

        List<string> EOLM11AV = new List<string>();
        List<string> EOLM11LSL = new List<string>();
        List<string> EOLM11USL = new List<string>();

        List<string> Scan3DMeasuring0AV = new List<string>();
        List<string> Scan3DMeasuring0LSL = new List<string>();
        List<string> Scan3DMeasuring0USL = new List<string>();

        List<string> Scan3DMeasuring1AV = new List<string>();
        List<string> Scan3DMeasuring1LSL = new List<string>();
        List<string> Scan3DMeasuring1USL = new List<string>();

        List<string> Scan3DMeasuring2AV = new List<string>();
        List<string> Scan3DMeasuring2LSL = new List<string>();
        List<string> Scan3DMeasuring2USL = new List<string>();

        List<string> Scan3DMeasuring3AV = new List<string>();
        List<string> Scan3DMeasuring3LSL = new List<string>();
        List<string> Scan3DMeasuring3USL = new List<string>();

        List<string> Scan3DMeasuring4AV = new List<string>();
        List<string> Scan3DMeasuring4LSL = new List<string>();
        List<string> Scan3DMeasuring4USL = new List<string>();

        List<string> Scan3DMeasuring5AV = new List<string>();
        List<string> Scan3DMeasuring5LSL = new List<string>();
        List<string> Scan3DMeasuring5USL = new List<string>();

        List<string> Scan3DMeasuring6AV = new List<string>();
        List<string> Scan3DMeasuring6LSL = new List<string>();
        List<string> Scan3DMeasuring6USL = new List<string>();

        List<string> Scan3DMeasuring7AV = new List<string>();
        List<string> Scan3DMeasuring7LSL = new List<string>();
        List<string> Scan3DMeasuring7USL = new List<string>();



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

            listOfLslUsl.Add(ProfileScanPoint0LSL);
            listOfLslUsl.Add(ProfileScanPoint0USL);
            listOfAV.Add(ProfileScanPoint0AV);

            listOfLslUsl.Add(ProfileScanPoint1LSL);
            listOfLslUsl.Add(ProfileScanPoint1USL);
            listOfAV.Add(ProfileScanPoint1AV);

            listOfLslUsl.Add(ProfileScanPoint2LSL);
            listOfLslUsl.Add(ProfileScanPoint2USL);
            listOfAV.Add(ProfileScanPoint2AV);

            listOfLslUsl.Add(ProfileScanPoint3LSL);
            listOfLslUsl.Add(ProfileScanPoint3USL);
            listOfAV.Add(ProfileScanPoint3AV);

            listOfLslUsl.Add(ProfileScanPoint4LSL);
            listOfLslUsl.Add(ProfileScanPoint4USL);
            listOfAV.Add(ProfileScanPoint4AV);

            listOfLslUsl.Add(ProfileScanPoint5LSL);
            listOfLslUsl.Add(ProfileScanPoint5USL);
            listOfAV.Add(ProfileScanPoint5AV);

            listOfLslUsl.Add(JoiningForceLSL);
            listOfLslUsl.Add(JoiningForceUSL);
            listOfAV.Add(JoiningForceAV);

            listOfLslUsl.Add(LaserWeldingSettingDistanceLSL);
            listOfLslUsl.Add(LaserWeldingSettingDistanceUSL);
            listOfAV.Add(LaserWeldingSettingDistanceAV);

            listOfLslUsl.Add(LaserWeldingStartPositionLSL);
            listOfLslUsl.Add(LaserWeldingStartPositionUSL);
            listOfAV.Add(LaserWeldingStartPositionAV);

            listOfLslUsl.Add(LaserWeldingShutdownDistanceLSL);
            listOfLslUsl.Add(LaserWeldingShutdownDistanceUSL);
            listOfAV.Add(LaserWeldingShutdownDistanceAV);

            listOfLslUsl.Add(LaserWeldingWeldingtimeLSL);
            listOfLslUsl.Add(LaserWeldingWeldingtimeUSL);
            listOfAV.Add(LaserWeldingWeldingtimeAV);

            listOfLslUsl.Add(ForceLSL);
            listOfLslUsl.Add(ForceUSL);
            listOfAV.Add(ForceAV);

            listOfLslUsl.Add(LeaktestStartPressureLSL);
            listOfLslUsl.Add(LeaktestStartPressureUSL);
            listOfAV.Add(LeaktestStartPressureAV);

            listOfLslUsl.Add(LeaktestLeakageLSL);
            listOfLslUsl.Add(LeaktestLeakageUSL);
            listOfAV.Add(LeaktestLeakageAV);

            listOfLslUsl.Add(EOLM01LSL);
            listOfLslUsl.Add(EOLM01USL);
            listOfAV.Add(EOLM01AV);

            listOfLslUsl.Add(EOLM02LSL);
            listOfLslUsl.Add(EOLM02USL);
            listOfAV.Add(EOLM02AV);

            listOfLslUsl.Add(EOLM03LSL);
            listOfLslUsl.Add(EOLM03USL);
            listOfAV.Add(EOLM03AV);

            listOfLslUsl.Add(EOLM04LSL);
            listOfLslUsl.Add(EOLM04USL);
            listOfAV.Add(EOLM04AV);

            listOfLslUsl.Add(EOLM05LSL);
            listOfLslUsl.Add(EOLM05USL);
            listOfAV.Add(EOLM05AV);

            listOfLslUsl.Add(EOLM06LSL);
            listOfLslUsl.Add(EOLM06USL);
            listOfAV.Add(EOLM06AV);

            listOfLslUsl.Add(EOLM07LSL);
            listOfLslUsl.Add(EOLM07USL);
            listOfAV.Add(EOLM07AV);

            listOfLslUsl.Add(EOLM08LSL);
            listOfLslUsl.Add(EOLM08USL);
            listOfAV.Add(EOLM08AV);

            listOfLslUsl.Add(EOLM09LSL);
            listOfLslUsl.Add(EOLM09USL);
            listOfAV.Add(EOLM09AV);

            listOfLslUsl.Add(EOLM10LSL);
            listOfLslUsl.Add(EOLM10USL);
            listOfAV.Add(EOLM10AV);

            listOfLslUsl.Add(EOLM11LSL);
            listOfLslUsl.Add(EOLM11USL);
            listOfAV.Add(EOLM11AV);

            listOfLslUsl.Add(Scan3DMeasuring0LSL);
            listOfLslUsl.Add(Scan3DMeasuring0USL);
            listOfAV.Add(Scan3DMeasuring0AV);

            listOfLslUsl.Add(Scan3DMeasuring1LSL);
            listOfLslUsl.Add(Scan3DMeasuring1USL);
            listOfAV.Add(Scan3DMeasuring1AV);

            listOfLslUsl.Add(Scan3DMeasuring2LSL);
            listOfLslUsl.Add(Scan3DMeasuring2USL);
            listOfAV.Add(Scan3DMeasuring2AV);

            listOfLslUsl.Add(Scan3DMeasuring3LSL);
            listOfLslUsl.Add(Scan3DMeasuring3USL);
            listOfAV.Add(Scan3DMeasuring3AV);

            listOfLslUsl.Add(Scan3DMeasuring4LSL);
            listOfLslUsl.Add(Scan3DMeasuring4USL);
            listOfAV.Add(Scan3DMeasuring4AV);

            listOfLslUsl.Add(Scan3DMeasuring5LSL);
            listOfLslUsl.Add(Scan3DMeasuring5USL);
            listOfAV.Add(Scan3DMeasuring5AV);

            listOfLslUsl.Add(Scan3DMeasuring6LSL);
            listOfLslUsl.Add(Scan3DMeasuring6USL);
            listOfAV.Add(Scan3DMeasuring6AV);

            listOfLslUsl.Add(Scan3DMeasuring7LSL);
            listOfLslUsl.Add(Scan3DMeasuring7USL);
            listOfAV.Add(Scan3DMeasuring7AV);


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

            txt_rangeValLSL.Add(txt_rangeProfileScanPoint0ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeProfileScanPoint0ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeProfileScanPoint1ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeProfileScanPoint1ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeProfileScanPoint2ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeProfileScanPoint2ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeProfileScanPoint3ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeProfileScanPoint3ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeProfileScanPoint4ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeProfileScanPoint4ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeProfileScanPoint5ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeProfileScanPoint5ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeJoiningForceValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeJoiningForceValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeLaserWeldingSettingDistanceValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeLaserWeldingSettingDistanceValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeLaserWeldingStartPositionValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeLaserWeldingStartPositionValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeLaserWeldingShutdownDistanceValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeLaserWeldingShutdownDistanceValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeLaserWeldingWeldingtimeValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeLaserWeldingWeldingtimeValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeForceValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeForceValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeLeaktestStartPressureValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeLeaktestStartPressureValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeLeaktestLeakageValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeLeaktestLeakageValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM01ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM01ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM02ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM02ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM03ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM03ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM04ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM04ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM05ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM05ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM06ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM06ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM07ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM07ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM08ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM08ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM09ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM09ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM10ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM10ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeEOLM11ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeEOLM11ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring0ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring0ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring1ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring1ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring2ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring2ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring3ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring3ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring4ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring4ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring5ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring5ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring6ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring6ValUSL.Text);

            txt_rangeValLSL.Add(txt_rangeScan3DMeasuring7ValLSL.Text);
            txt_rangeValUSL.Add(txt_rangeScan3DMeasuring7ValUSL.Text);


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
            lbl_status.Add(lbl_statusProfileScanPoint0AV);
            lbl_status.Add(lbl_statusProfileScanPoint1AV);
            lbl_status.Add(lbl_statusProfileScanPoint2AV);
            lbl_status.Add(lbl_statusProfileScanPoint3AV);
            lbl_status.Add(lbl_statusProfileScanPoint4AV);
            lbl_status.Add(lbl_statusProfileScanPoint5AV);
            lbl_status.Add(lbl_statusJoiningForceAV);
            lbl_status.Add(lbl_statusLaserWeldingSettingDistanceAV);
            lbl_status.Add(lbl_statusLaserWeldingStartPositionAV);
            lbl_status.Add(lbl_statusLaserWeldingShutdownDistanceAV);
            lbl_status.Add(lbl_statusLaserWeldingWeldingtimeAV);
            lbl_status.Add(lbl_statusForceAV);
            lbl_status.Add(lbl_statusLeaktestStartPressureAV);
            lbl_status.Add(lbl_statusLeaktestLeakageAV);
            lbl_status.Add(lbl_statusEOLM01AV);
            lbl_status.Add(lbl_statusEOLM02AV);
            lbl_status.Add(lbl_statusEOLM03AV);
            lbl_status.Add(lbl_statusEOLM04AV);
            lbl_status.Add(lbl_statusEOLM05AV);
            lbl_status.Add(lbl_statusEOLM06AV);
            lbl_status.Add(lbl_statusEOLM07AV);
            lbl_status.Add(lbl_statusEOLM08AV);
            lbl_status.Add(lbl_statusEOLM09AV);
            lbl_status.Add(lbl_statusEOLM10AV);
            lbl_status.Add(lbl_statusEOLM11AV);
            lbl_status.Add(lbl_statusScan3DMeasuring0AV);
            lbl_status.Add(lbl_statusScan3DMeasuring1AV);
            lbl_status.Add(lbl_statusScan3DMeasuring2AV);
            lbl_status.Add(lbl_statusScan3DMeasuring3AV);
            lbl_status.Add(lbl_statusScan3DMeasuring4AV);
            lbl_status.Add(lbl_statusScan3DMeasuring5AV);
            lbl_status.Add(lbl_statusScan3DMeasuring6AV);
            lbl_status.Add(lbl_statusScan3DMeasuring7AV);


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
            lbl_statusBackColor.Add(lbl_statusProfileScanPoint0AV);
            lbl_statusBackColor.Add(lbl_statusProfileScanPoint1AV);
            lbl_statusBackColor.Add(lbl_statusProfileScanPoint2AV);
            lbl_statusBackColor.Add(lbl_statusProfileScanPoint3AV);
            lbl_statusBackColor.Add(lbl_statusProfileScanPoint4AV);
            lbl_statusBackColor.Add(lbl_statusProfileScanPoint5AV);
            lbl_statusBackColor.Add(lbl_statusJoiningForceAV);
            lbl_statusBackColor.Add(lbl_statusLaserWeldingSettingDistanceAV);
            lbl_statusBackColor.Add(lbl_statusLaserWeldingStartPositionAV);
            lbl_statusBackColor.Add(lbl_statusLaserWeldingShutdownDistanceAV);
            lbl_statusBackColor.Add(lbl_statusLaserWeldingWeldingtimeAV);
            lbl_statusBackColor.Add(lbl_statusForceAV);
            lbl_statusBackColor.Add(lbl_statusLeaktestStartPressureAV);
            lbl_statusBackColor.Add(lbl_statusLeaktestLeakageAV);
            lbl_statusBackColor.Add(lbl_statusEOLM01AV);
            lbl_statusBackColor.Add(lbl_statusEOLM02AV);
            lbl_statusBackColor.Add(lbl_statusEOLM03AV);
            lbl_statusBackColor.Add(lbl_statusEOLM04AV);
            lbl_statusBackColor.Add(lbl_statusEOLM05AV);
            lbl_statusBackColor.Add(lbl_statusEOLM06AV);
            lbl_statusBackColor.Add(lbl_statusEOLM07AV);
            lbl_statusBackColor.Add(lbl_statusEOLM08AV);
            lbl_statusBackColor.Add(lbl_statusEOLM09AV);
            lbl_statusBackColor.Add(lbl_statusEOLM10AV);
            lbl_statusBackColor.Add(lbl_statusEOLM11AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring0AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring1AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring2AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring3AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring4AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring5AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring6AV);
            lbl_statusBackColor.Add(lbl_statusScan3DMeasuring7AV);



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
            txt_rangeProfileScanPoint0ValAV.Text = ProfileScanPoint0AV[0];
            txt_rangeProfileScanPoint1ValAV.Text = ProfileScanPoint1AV[0];
            txt_rangeProfileScanPoint2ValAV.Text = ProfileScanPoint2AV[0];
            txt_rangeProfileScanPoint3ValAV.Text = ProfileScanPoint3AV[0];
            txt_rangeProfileScanPoint4ValAV.Text = ProfileScanPoint4AV[0];
            txt_rangeProfileScanPoint5ValAV.Text = ProfileScanPoint5AV[0];
            txt_rangeJoiningForceValAV.Text = JoiningForceAV[0];
            txt_rangeLaserWeldingSettingDistanceValAV.Text = LaserWeldingSettingDistanceAV[0];
            txt_rangeLaserWeldingStartPositionValAV.Text = LaserWeldingStartPositionAV[0];
            txt_rangeLaserWeldingShutdownDistanceValAV.Text = LaserWeldingShutdownDistanceAV[0];
            txt_rangeLaserWeldingWeldingtimeValAV.Text = LaserWeldingWeldingtimeAV[0];
            txt_rangeForceValAV.Text = ForceAV[0];
            txt_rangeLeaktestStartPressureValAV.Text = LeaktestStartPressureAV[0];
            txt_rangeLeaktestLeakageValAV.Text = LeaktestLeakageAV[0];
            txt_rangeEOLM01ValAV.Text = EOLM01AV[0];
            txt_rangeEOLM02ValAV.Text = EOLM02AV[0];
            txt_rangeEOLM03ValAV.Text = EOLM03AV[0];
            txt_rangeEOLM04ValAV.Text = EOLM04AV[0];
            txt_rangeEOLM05ValAV.Text = EOLM05AV[0];
            txt_rangeEOLM06ValAV.Text = EOLM06AV[0];
            txt_rangeEOLM07ValAV.Text = EOLM07AV[0];
            txt_rangeEOLM08ValAV.Text = EOLM08AV[0];
            txt_rangeEOLM09ValAV.Text = EOLM09AV[0];
            txt_rangeEOLM10ValAV.Text = EOLM10AV[0];
            txt_rangeEOLM11ValAV.Text = EOLM11AV[0];
            txt_rangeScan3DMeasuring0ValAV.Text = Scan3DMeasuring0AV[0];
            txt_rangeScan3DMeasuring1ValAV.Text = Scan3DMeasuring1AV[0];
            txt_rangeScan3DMeasuring2ValAV.Text = Scan3DMeasuring2AV[0];
            txt_rangeScan3DMeasuring3ValAV.Text = Scan3DMeasuring3AV[0];
            txt_rangeScan3DMeasuring4ValAV.Text = Scan3DMeasuring4AV[0];
            txt_rangeScan3DMeasuring5ValAV.Text = Scan3DMeasuring5AV[0];
            txt_rangeScan3DMeasuring6ValAV.Text = Scan3DMeasuring6AV[0];
            txt_rangeScan3DMeasuring7ValAV.Text = Scan3DMeasuring7AV[0];


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
            ComparareValoriProfileScanPoint0LSL();
            ComparareValoriProfileScanPoint0USL();

            ComparareValoriProfileScanPoint1LSL();
            ComparareValoriProfileScanPoint1USL();

            ComparareValoriProfileScanPoint2LSL();
            ComparareValoriProfileScanPoint2USL();

            ComparareValoriProfileScanPoint3LSL();
            ComparareValoriProfileScanPoint3USL();

            ComparareValoriProfileScanPoint4LSL();
            ComparareValoriProfileScanPoint4USL();

            ComparareValoriProfileScanPoint5LSL();
            ComparareValoriProfileScanPoint5USL();

            ComparareValoriJoiningForceLSL();
            ComparareValoriJoiningForceUSL();

            ComparareValoriLaserWeldingSettingDistanceLSL();
            ComparareValoriLaserWeldingSettingDistanceUSL();

            ComparareValoriLaserWeldingStartPositionLSL();
            ComparareValoriLaserWeldingStartPositionUSL();

            ComparareValoriLaserWeldingShutdownDistanceLSL();
            ComparareValoriLaserWeldingShutdownDistanceUSL();

            ComparareValoriLaserWeldingWeldingtimeLSL();
            ComparareValoriLaserWeldingWeldingtimeUSL();

            ComparareValoriForceLSL();
            ComparareValoriForceUSL();

            ComparareValoriLeaktestStartPressureLSL();
            ComparareValoriLeaktestStartPressureUSL();

            ComparareValoriLeaktestLeakageLSL();
            ComparareValoriLeaktestLeakageUSL();

            ComparareValoriEOLM01LSL();
            ComparareValoriEOLM01USL();

            ComparareValoriEOLM02LSL();
            ComparareValoriEOLM02USL();

            ComparareValoriEOLM03LSL();
            ComparareValoriEOLM03USL();

            ComparareValoriEOLM04LSL();
            ComparareValoriEOLM04USL();

            ComparareValoriEOLM05LSL();
            ComparareValoriEOLM05USL();

            ComparareValoriEOLM06LSL();
            ComparareValoriEOLM06USL();

            ComparareValoriEOLM07LSL();
            ComparareValoriEOLM07USL();

            ComparareValoriEOLM08LSL();
            ComparareValoriEOLM08USL();

            ComparareValoriEOLM09LSL();
            ComparareValoriEOLM09USL();

            ComparareValoriEOLM10LSL();
            ComparareValoriEOLM10USL();

            ComparareValoriEOLM11LSL();
            ComparareValoriEOLM11USL();


            ComparareValoriScan3DMeasuring0LSL();
            ComparareValoriScan3DMeasuring0USL();

            ComparareValoriScan3DMeasuring1LSL();
            ComparareValoriScan3DMeasuring1USL();

            ComparareValoriScan3DMeasuring2LSL();
            ComparareValoriScan3DMeasuring2USL();

            ComparareValoriScan3DMeasuring3LSL();
            ComparareValoriScan3DMeasuring3USL();

            ComparareValoriScan3DMeasuring4LSL();
            ComparareValoriScan3DMeasuring4USL();

            ComparareValoriScan3DMeasuring5LSL();
            ComparareValoriScan3DMeasuring5USL();

            ComparareValoriScan3DMeasuring6LSL();
            ComparareValoriScan3DMeasuring6USL();

            ComparareValoriScan3DMeasuring7LSL();
            ComparareValoriScan3DMeasuring7USL();

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

            string ProfileScanPoint0AVRange = Properties.Settings.Default.txt_rangeProfileScanPoint0 + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint0.Text = Properties.Settings.Default.txt_rangeProfileScanPoint0 + txt_mediacoloana.Text;
            string ProfileScanPoint0LSLRange = Properties.Settings.Default.txt_rangeProfileScanPoint0LSL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint0LSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint0LSL + txt_mediacoloana.Text;
            string ProfileScanPoint0USLRange = Properties.Settings.Default.txt_rangeProfileScanPoint0USL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint0USL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint0USL + txt_mediacoloana.Text;

            string ProfileScanPoint1AVRange = Properties.Settings.Default.txt_rangeProfileScanPoint1 + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint1.Text = Properties.Settings.Default.txt_rangeProfileScanPoint1 + txt_mediacoloana.Text;
            string ProfileScanPoint1LSLRange = Properties.Settings.Default.txt_rangeProfileScanPoint1LSL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint1LSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint1LSL + txt_mediacoloana.Text;
            string ProfileScanPoint1USLRange = Properties.Settings.Default.txt_rangeProfileScanPoint1USL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint1USL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint1USL + txt_mediacoloana.Text;

            string ProfileScanPoint2AVRange = Properties.Settings.Default.txt_rangeProfileScanPoint2 + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint2.Text = Properties.Settings.Default.txt_rangeProfileScanPoint2 + txt_mediacoloana.Text;
            string ProfileScanPoint2LSLRange = Properties.Settings.Default.txt_rangeProfileScanPoint2LSL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint2LSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint2LSL + txt_mediacoloana.Text;
            string ProfileScanPoint2USLRange = Properties.Settings.Default.txt_rangeProfileScanPoint2USL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint2USL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint2USL + txt_mediacoloana.Text;

            string ProfileScanPoint3AVRange = Properties.Settings.Default.txt_rangeProfileScanPoint3 + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint3.Text = Properties.Settings.Default.txt_rangeProfileScanPoint3 + txt_mediacoloana.Text;
            string ProfileScanPoint3LSLRange = Properties.Settings.Default.txt_rangeProfileScanPoint3LSL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint3LSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint3LSL + txt_mediacoloana.Text;
            string ProfileScanPoint3USLRange = Properties.Settings.Default.txt_rangeProfileScanPoint3USL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint3USL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint3USL + txt_mediacoloana.Text;

            string ProfileScanPoint4AVRange = Properties.Settings.Default.txt_rangeProfileScanPoint4 + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint4.Text = Properties.Settings.Default.txt_rangeProfileScanPoint4 + txt_mediacoloana.Text;
            string ProfileScanPoint4LSLRange = Properties.Settings.Default.txt_rangeProfileScanPoint4LSL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint4LSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint4LSL + txt_mediacoloana.Text;
            string ProfileScanPoint4USLRange = Properties.Settings.Default.txt_rangeProfileScanPoint4USL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint4USL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint4USL + txt_mediacoloana.Text;

            string ProfileScanPoint5AVRange = Properties.Settings.Default.txt_rangeProfileScanPoint5 + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint5.Text = Properties.Settings.Default.txt_rangeProfileScanPoint5 + txt_mediacoloana.Text;
            string ProfileScanPoint5LSLRange = Properties.Settings.Default.txt_rangeProfileScanPoint5LSL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint5LSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint5LSL + txt_mediacoloana.Text;
            string ProfileScanPoint5USLRange = Properties.Settings.Default.txt_rangeProfileScanPoint5USL + txt_mediacoloana.Text;
            txt_rangeProfileScanPoint5USL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint5USL + txt_mediacoloana.Text;

            string JoiningForceAVRange = Properties.Settings.Default.txt_rangeJoiningForce + txt_mediacoloana.Text;
            txt_rangeJoiningForce.Text = Properties.Settings.Default.txt_rangeJoiningForce + txt_mediacoloana.Text;
            string JoiningForceLSLRange = Properties.Settings.Default.txt_rangeJoiningForceLSL + txt_mediacoloana.Text;
            txt_rangeJoiningForceLSL.Text = Properties.Settings.Default.txt_rangeJoiningForceLSL + txt_mediacoloana.Text;
            string JoiningForceUSLRange = Properties.Settings.Default.txt_rangeJoiningForceUSL + txt_mediacoloana.Text;
            txt_rangeJoiningForceUSL.Text = Properties.Settings.Default.txt_rangeJoiningForceUSL + txt_mediacoloana.Text;

            string LaserWeldingSettingDistanceAVRange = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistance + txt_mediacoloana.Text;
            txt_rangeLaserWeldingSettingDistance.Text = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistance + txt_mediacoloana.Text;
            string LaserWeldingSettingDistanceLSLRange = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistanceLSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingSettingDistanceLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistanceLSL + txt_mediacoloana.Text;
            string LaserWeldingSettingDistanceUSLRange = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistanceUSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingSettingDistanceUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistanceUSL + txt_mediacoloana.Text;

            string LaserWeldingStartPositionAVRange = Properties.Settings.Default.txt_rangeLaserWeldingStartPosition + txt_mediacoloana.Text;
            txt_rangeLaserWeldingStartPosition.Text = Properties.Settings.Default.txt_rangeLaserWeldingStartPosition + txt_mediacoloana.Text;
            string LaserWeldingStartPositionLSLRange = Properties.Settings.Default.txt_rangeLaserWeldingStartPositionLSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingStartPositionLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingStartPositionLSL + txt_mediacoloana.Text;
            string LaserWeldingStartPositionUSLRange = Properties.Settings.Default.txt_rangeLaserWeldingStartPositionUSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingStartPositionUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingStartPositionUSL + txt_mediacoloana.Text;

            string LaserWeldingShutdownDistanceAVRange = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistance + txt_mediacoloana.Text;
            txt_rangeLaserWeldingShutdownDistance.Text = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistance + txt_mediacoloana.Text;
            string LaserWeldingShutdownDistanceLSLRange = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistanceLSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingShutdownDistanceLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistanceLSL + txt_mediacoloana.Text;
            string LaserWeldingShutdownDistanceUSLRange = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistanceUSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingShutdownDistanceUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistanceUSL + txt_mediacoloana.Text;

            string LaserWeldingWeldingtimeAVRange = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtime + txt_mediacoloana.Text;
            txt_rangeLaserWeldingWeldingtime.Text = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtime + txt_mediacoloana.Text;
            string LaserWeldingWeldingtimeLSLRange = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtimeLSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingWeldingtimeLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtimeLSL + txt_mediacoloana.Text;
            string LaserWeldingWeldingtimeUSLRange = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtimeUSL + txt_mediacoloana.Text;
            txt_rangeLaserWeldingWeldingtimeUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtimeUSL + txt_mediacoloana.Text;

            string ForceAVRange = Properties.Settings.Default.txt_rangeForce + txt_mediacoloana.Text;
            txt_rangeForce.Text = Properties.Settings.Default.txt_rangeForce + txt_mediacoloana.Text;
            string ForceLSLRange = Properties.Settings.Default.txt_rangeForceLSL + txt_mediacoloana.Text;
            txt_rangeForceLSL.Text = Properties.Settings.Default.txt_rangeForceLSL + txt_mediacoloana.Text;
            string ForceUSLRange = Properties.Settings.Default.txt_rangeForceUSL + txt_mediacoloana.Text;
            txt_rangeForceUSL.Text = Properties.Settings.Default.txt_rangeForceUSL + txt_mediacoloana.Text;

            string LeaktestStartPressureAVRange = Properties.Settings.Default.txt_rangeLeaktestStartPressure + txt_mediacoloana.Text;
            txt_rangeLeaktestStartPressure.Text = Properties.Settings.Default.txt_rangeLeaktestStartPressure + txt_mediacoloana.Text;
            string LeaktestStartPressureLSLRange = Properties.Settings.Default.txt_rangeLeaktestStartPressureLSL + txt_mediacoloana.Text;
            txt_rangeLeaktestStartPressureLSL.Text = Properties.Settings.Default.txt_rangeLeaktestStartPressureLSL + txt_mediacoloana.Text;
            string LeaktestStartPressureUSLRange = Properties.Settings.Default.txt_rangeLeaktestStartPressureUSL + txt_mediacoloana.Text;
            txt_rangeLeaktestStartPressureUSL.Text = Properties.Settings.Default.txt_rangeLeaktestStartPressureUSL + txt_mediacoloana.Text;

            string LeaktestLeakageAVRange = Properties.Settings.Default.txt_rangeLeaktestLeakage + txt_mediacoloana.Text;
            txt_rangeLeaktestLeakage.Text = Properties.Settings.Default.txt_rangeLeaktestLeakage + txt_mediacoloana.Text;
            string LeaktestLeakageLSLRange = Properties.Settings.Default.txt_rangeLeaktestLeakageLSL + txt_mediacoloana.Text;
            txt_rangeLeaktestLeakageLSL.Text = Properties.Settings.Default.txt_rangeLeaktestLeakageLSL + txt_mediacoloana.Text;
            string LeaktestLeakageUSLRange = Properties.Settings.Default.txt_rangeLeaktestLeakageUSL + txt_mediacoloana.Text;
            txt_rangeLeaktestLeakageUSL.Text = Properties.Settings.Default.txt_rangeLeaktestLeakageUSL + txt_mediacoloana.Text;

            string EOLM01AVRange = Properties.Settings.Default.txt_rangeEOLM01 + txt_mediacoloana.Text;
            txt_rangeEOLM01.Text = Properties.Settings.Default.txt_rangeEOLM01 + txt_mediacoloana.Text;
            string EOLM01LSLRange = Properties.Settings.Default.txt_rangeEOLM01LSL + txt_mediacoloana.Text;
            txt_rangeEOLM01LSL.Text = Properties.Settings.Default.txt_rangeEOLM01LSL + txt_mediacoloana.Text;
            string EOLM01USLRange = Properties.Settings.Default.txt_rangeEOLM01USL + txt_mediacoloana.Text;
            txt_rangeEOLM01USL.Text = Properties.Settings.Default.txt_rangeEOLM01USL + txt_mediacoloana.Text;

            string EOLM02AVRange = Properties.Settings.Default.txt_rangeEOLM02 + txt_mediacoloana.Text;
            txt_rangeEOLM02.Text = Properties.Settings.Default.txt_rangeEOLM02 + txt_mediacoloana.Text;
            string EOLM02LSLRange = Properties.Settings.Default.txt_rangeEOLM02LSL + txt_mediacoloana.Text;
            txt_rangeEOLM02LSL.Text = Properties.Settings.Default.txt_rangeEOLM02LSL + txt_mediacoloana.Text;
            string EOLM02USLRange = Properties.Settings.Default.txt_rangeEOLM02USL + txt_mediacoloana.Text;
            txt_rangeEOLM02USL.Text = Properties.Settings.Default.txt_rangeEOLM02USL + txt_mediacoloana.Text;

            string EOLM03AVRange = Properties.Settings.Default.txt_rangeEOLM03 + txt_mediacoloana.Text;
            txt_rangeEOLM03.Text = Properties.Settings.Default.txt_rangeEOLM03 + txt_mediacoloana.Text;
            string EOLM03LSLRange = Properties.Settings.Default.txt_rangeEOLM03LSL + txt_mediacoloana.Text;
            txt_rangeEOLM03LSL.Text = Properties.Settings.Default.txt_rangeEOLM03LSL + txt_mediacoloana.Text;
            string EOLM03USLRange = Properties.Settings.Default.txt_rangeEOLM03USL + txt_mediacoloana.Text;
            txt_rangeEOLM03USL.Text = Properties.Settings.Default.txt_rangeEOLM03USL + txt_mediacoloana.Text;

            string EOLM04AVRange = Properties.Settings.Default.txt_rangeEOLM04 + txt_mediacoloana.Text;
            txt_rangeEOLM04.Text = Properties.Settings.Default.txt_rangeEOLM04 + txt_mediacoloana.Text;
            string EOLM04LSLRange = Properties.Settings.Default.txt_rangeEOLM04LSL + txt_mediacoloana.Text;
            txt_rangeEOLM04LSL.Text = Properties.Settings.Default.txt_rangeEOLM04LSL + txt_mediacoloana.Text;
            string EOLM04USLRange = Properties.Settings.Default.txt_rangeEOLM04USL + txt_mediacoloana.Text;
            txt_rangeEOLM04USL.Text = Properties.Settings.Default.txt_rangeEOLM04USL + txt_mediacoloana.Text;

            string EOLM05AVRange = Properties.Settings.Default.txt_rangeEOLM05 + txt_mediacoloana.Text;
            txt_rangeEOLM05.Text = Properties.Settings.Default.txt_rangeEOLM05 + txt_mediacoloana.Text;
            string EOLM05LSLRange = Properties.Settings.Default.txt_rangeEOLM05LSL + txt_mediacoloana.Text;
            txt_rangeEOLM05LSL.Text = Properties.Settings.Default.txt_rangeEOLM05LSL + txt_mediacoloana.Text;
            string EOLM05USLRange = Properties.Settings.Default.txt_rangeEOLM05USL + txt_mediacoloana.Text;
            txt_rangeEOLM05USL.Text = Properties.Settings.Default.txt_rangeEOLM05USL + txt_mediacoloana.Text;

            string EOLM06AVRange = Properties.Settings.Default.txt_rangeEOLM06 + txt_mediacoloana.Text;
            txt_rangeEOLM06.Text = Properties.Settings.Default.txt_rangeEOLM06 + txt_mediacoloana.Text;
            string EOLM06LSLRange = Properties.Settings.Default.txt_rangeEOLM06LSL + txt_mediacoloana.Text;
            txt_rangeEOLM06LSL.Text = Properties.Settings.Default.txt_rangeEOLM06LSL + txt_mediacoloana.Text;
            string EOLM06USLRange = Properties.Settings.Default.txt_rangeEOLM06USL + txt_mediacoloana.Text;
            txt_rangeEOLM06USL.Text = Properties.Settings.Default.txt_rangeEOLM06USL + txt_mediacoloana.Text;

            string EOLM07AVRange = Properties.Settings.Default.txt_rangeEOLM07 + txt_mediacoloana.Text;
            txt_rangeEOLM07.Text = Properties.Settings.Default.txt_rangeEOLM07 + txt_mediacoloana.Text;
            string EOLM07LSLRange = Properties.Settings.Default.txt_rangeEOLM07LSL + txt_mediacoloana.Text;
            txt_rangeEOLM07LSL.Text = Properties.Settings.Default.txt_rangeEOLM07LSL + txt_mediacoloana.Text;
            string EOLM07USLRange = Properties.Settings.Default.txt_rangeEOLM07USL + txt_mediacoloana.Text;
            txt_rangeEOLM07USL.Text = Properties.Settings.Default.txt_rangeEOLM07USL + txt_mediacoloana.Text;

            string EOLM08AVRange = Properties.Settings.Default.txt_rangeEOLM08 + txt_mediacoloana.Text;
            txt_rangeEOLM08.Text = Properties.Settings.Default.txt_rangeEOLM08 + txt_mediacoloana.Text;
            string EOLM08LSLRange = Properties.Settings.Default.txt_rangeEOLM08LSL + txt_mediacoloana.Text;
            txt_rangeEOLM08LSL.Text = Properties.Settings.Default.txt_rangeEOLM08LSL + txt_mediacoloana.Text;
            string EOLM08USLRange = Properties.Settings.Default.txt_rangeEOLM08USL + txt_mediacoloana.Text;
            txt_rangeEOLM08USL.Text = Properties.Settings.Default.txt_rangeEOLM08USL + txt_mediacoloana.Text;

            string EOLM09AVRange = Properties.Settings.Default.txt_rangeEOLM09 + txt_mediacoloana.Text;
            txt_rangeEOLM09.Text = Properties.Settings.Default.txt_rangeEOLM09 + txt_mediacoloana.Text;
            string EOLM09LSLRange = Properties.Settings.Default.txt_rangeEOLM09LSL + txt_mediacoloana.Text;
            txt_rangeEOLM09LSL.Text = Properties.Settings.Default.txt_rangeEOLM09LSL + txt_mediacoloana.Text;
            string EOLM09USLRange = Properties.Settings.Default.txt_rangeEOLM09USL + txt_mediacoloana.Text;
            txt_rangeEOLM09USL.Text = Properties.Settings.Default.txt_rangeEOLM09USL + txt_mediacoloana.Text;

            string EOLM10AVRange = Properties.Settings.Default.txt_rangeEOLM10 + txt_mediacoloana.Text;
            txt_rangeEOLM10.Text = Properties.Settings.Default.txt_rangeEOLM10 + txt_mediacoloana.Text;
            string EOLM10LSLRange = Properties.Settings.Default.txt_rangeEOLM10LSL + txt_mediacoloana.Text;
            txt_rangeEOLM10LSL.Text = Properties.Settings.Default.txt_rangeEOLM10LSL + txt_mediacoloana.Text;
            string EOLM10USLRange = Properties.Settings.Default.txt_rangeEOLM10USL + txt_mediacoloana.Text;
            txt_rangeEOLM10USL.Text = Properties.Settings.Default.txt_rangeEOLM10USL + txt_mediacoloana.Text;

            string EOLM11AVRange = Properties.Settings.Default.txt_rangeEOLM11 + txt_mediacoloana.Text;
            txt_rangeEOLM11.Text = Properties.Settings.Default.txt_rangeEOLM11 + txt_mediacoloana.Text;
            string EOLM11LSLRange = Properties.Settings.Default.txt_rangeEOLM11LSL + txt_mediacoloana.Text;
            txt_rangeEOLM11LSL.Text = Properties.Settings.Default.txt_rangeEOLM11LSL + txt_mediacoloana.Text;
            string EOLM11USLRange = Properties.Settings.Default.txt_rangeEOLM11USL + txt_mediacoloana.Text;
            txt_rangeEOLM11USL.Text = Properties.Settings.Default.txt_rangeEOLM11USL + txt_mediacoloana.Text;

            string Scan3DMeasuring0AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring0 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring0.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring0 + txt_mediacoloana.Text;
            string Scan3DMeasuring0LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring0LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring0LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring0LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring0USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring0USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring0USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring0USL + txt_mediacoloana.Text;

            string Scan3DMeasuring1AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring1 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring1.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring1 + txt_mediacoloana.Text;
            string Scan3DMeasuring1LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring1LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring1LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring1LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring1USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring1USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring1USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring1USL + txt_mediacoloana.Text;

            string Scan3DMeasuring2AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring2 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring2.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring2 + txt_mediacoloana.Text;
            string Scan3DMeasuring2LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring2LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring2LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring2LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring2USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring2USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring2USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring2USL + txt_mediacoloana.Text;

            string Scan3DMeasuring3AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring3 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring3.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring3 + txt_mediacoloana.Text;
            string Scan3DMeasuring3LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring3LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring3LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring3LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring3USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring3USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring3USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring3USL + txt_mediacoloana.Text;

            string Scan3DMeasuring4AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring4 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring4.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring4 + txt_mediacoloana.Text;
            string Scan3DMeasuring4LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring4LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring4LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring4LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring4USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring4USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring4USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring4USL + txt_mediacoloana.Text;

            string Scan3DMeasuring5AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring5 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring5.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring5 + txt_mediacoloana.Text;
            string Scan3DMeasuring5LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring5LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring5LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring5LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring5USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring5USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring5USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring5USL + txt_mediacoloana.Text;

            string Scan3DMeasuring6AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring6 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring6.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring6 + txt_mediacoloana.Text;
            string Scan3DMeasuring6LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring6LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring6LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring6LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring6USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring6USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring6USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring6USL + txt_mediacoloana.Text;

            string Scan3DMeasuring7AVRange = Properties.Settings.Default.txt_rangeScan3DMeasuring7 + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring7.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring7 + txt_mediacoloana.Text;
            string Scan3DMeasuring7LSLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring7LSL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring7LSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring7LSL + txt_mediacoloana.Text;
            string Scan3DMeasuring7USLRange = Properties.Settings.Default.txt_rangeScan3DMeasuring7USL + txt_mediacoloana.Text;
            txt_rangeScan3DMeasuring7USL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring7USL + txt_mediacoloana.Text;



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

            // ProfileScanPoint0 av range
            rangeDeCititAV.Add(ProfileScanPoint0AVRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint0LSLRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint0USLRange);

            // ProfileScanPoint1 av range
            rangeDeCititAV.Add(ProfileScanPoint1AVRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint1LSLRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint1USLRange);

            // ProfileScanPoint2 av range
            rangeDeCititAV.Add(ProfileScanPoint2AVRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint2LSLRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint2USLRange);

            // ProfileScanPoint3 av range
            rangeDeCititAV.Add(ProfileScanPoint3AVRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint3LSLRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint3USLRange);

            // ProfileScanPoint4 av range
            rangeDeCititAV.Add(ProfileScanPoint4AVRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint4LSLRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint4USLRange);

            // ProfileScanPoint5 av range
            rangeDeCititAV.Add(ProfileScanPoint5AVRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint5LSLRange);
            rangeDeCititLslUsl.Add(ProfileScanPoint5USLRange);

            // JoiningForce av range
            rangeDeCititAV.Add(JoiningForceAVRange);
            rangeDeCititLslUsl.Add(JoiningForceLSLRange);
            rangeDeCititLslUsl.Add(JoiningForceUSLRange);

            // LaserWeldingSettingDistance av range
            rangeDeCititAV.Add(LaserWeldingSettingDistanceAVRange);
            rangeDeCititLslUsl.Add(LaserWeldingSettingDistanceLSLRange);
            rangeDeCititLslUsl.Add(LaserWeldingSettingDistanceUSLRange);

            // LaserWeldingStartPosition av range
            rangeDeCititAV.Add(LaserWeldingStartPositionAVRange);
            rangeDeCititLslUsl.Add(LaserWeldingStartPositionLSLRange);
            rangeDeCititLslUsl.Add(LaserWeldingStartPositionUSLRange);

            // LaserWeldingShutdownDistance av range
            rangeDeCititAV.Add(LaserWeldingShutdownDistanceAVRange);
            rangeDeCititLslUsl.Add(LaserWeldingShutdownDistanceLSLRange);
            rangeDeCititLslUsl.Add(LaserWeldingShutdownDistanceUSLRange);

            // LaserWeldingWeldingtime av range
            rangeDeCititAV.Add(LaserWeldingWeldingtimeAVRange);
            rangeDeCititLslUsl.Add(LaserWeldingWeldingtimeLSLRange);
            rangeDeCititLslUsl.Add(LaserWeldingWeldingtimeUSLRange);


            // Force av range
            rangeDeCititAV.Add(ForceAVRange);
            rangeDeCititLslUsl.Add(ForceLSLRange);
            rangeDeCititLslUsl.Add(ForceUSLRange);

            // LeaktestStartPressure av range
            rangeDeCititAV.Add(LeaktestStartPressureAVRange);
            rangeDeCititLslUsl.Add(LeaktestStartPressureLSLRange);
            rangeDeCititLslUsl.Add(LeaktestStartPressureUSLRange);

            // LeaktestLeakage av range
            rangeDeCititAV.Add(LeaktestLeakageAVRange);
            rangeDeCititLslUsl.Add(LeaktestLeakageLSLRange);
            rangeDeCititLslUsl.Add(LeaktestLeakageUSLRange);

            // EOLM01 av range
            rangeDeCititAV.Add(EOLM01AVRange);
            rangeDeCititLslUsl.Add(EOLM01LSLRange);
            rangeDeCititLslUsl.Add(EOLM01USLRange);

            // EOLM02 av range
            rangeDeCititAV.Add(EOLM02AVRange);
            rangeDeCititLslUsl.Add(EOLM02LSLRange);
            rangeDeCititLslUsl.Add(EOLM02USLRange);

            // EOLM03 av range
            rangeDeCititAV.Add(EOLM03AVRange);
            rangeDeCititLslUsl.Add(EOLM03LSLRange);
            rangeDeCititLslUsl.Add(EOLM03USLRange);


            // EOLM04 av range
            rangeDeCititAV.Add(EOLM04AVRange);
            rangeDeCititLslUsl.Add(EOLM04LSLRange);
            rangeDeCititLslUsl.Add(EOLM04USLRange);


            // EOLM05 av range
            rangeDeCititAV.Add(EOLM05AVRange);
            rangeDeCititLslUsl.Add(EOLM05LSLRange);
            rangeDeCititLslUsl.Add(EOLM05USLRange);

            // EOLM06 av range
            rangeDeCititAV.Add(EOLM06AVRange);
            rangeDeCititLslUsl.Add(EOLM06LSLRange);
            rangeDeCititLslUsl.Add(EOLM06USLRange);


            // EOLM07 av range
            rangeDeCititAV.Add(EOLM07AVRange);
            rangeDeCititLslUsl.Add(EOLM07LSLRange);
            rangeDeCititLslUsl.Add(EOLM07USLRange);


            // EOLM08 av range
            rangeDeCititAV.Add(EOLM08AVRange);
            rangeDeCititLslUsl.Add(EOLM08LSLRange);
            rangeDeCititLslUsl.Add(EOLM08USLRange);


            // EOLM09 av range
            rangeDeCititAV.Add(EOLM09AVRange);
            rangeDeCititLslUsl.Add(EOLM09LSLRange);
            rangeDeCititLslUsl.Add(EOLM09USLRange);


            // EOLM10 av range
            rangeDeCititAV.Add(EOLM10AVRange);
            rangeDeCititLslUsl.Add(EOLM10LSLRange);
            rangeDeCititLslUsl.Add(EOLM10USLRange);


            // EOLM11 av range
            rangeDeCititAV.Add(EOLM11AVRange);
            rangeDeCititLslUsl.Add(EOLM11LSLRange);
            rangeDeCititLslUsl.Add(EOLM11USLRange);

            // Scan3DMeasuring0 av range
            rangeDeCititAV.Add(Scan3DMeasuring0AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring0LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring0USLRange);

            // Scan3DMeasuring1 av range
            rangeDeCititAV.Add(Scan3DMeasuring1AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring1LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring1USLRange);

            // Scan3DMeasuring2 av range
            rangeDeCititAV.Add(Scan3DMeasuring2AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring2LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring2USLRange);

            // Scan3DMeasuring3 av range
            rangeDeCititAV.Add(Scan3DMeasuring3AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring3LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring3USLRange);

            // Scan3DMeasuring4 av range
            rangeDeCititAV.Add(Scan3DMeasuring4AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring4LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring4USLRange);


            // Scan3DMeasuring5 av range
            rangeDeCititAV.Add(Scan3DMeasuring5AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring5LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring5USLRange);


            // Scan3DMeasuring6 av range
            rangeDeCititAV.Add(Scan3DMeasuring6AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring6LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring6USLRange);

            // Scan3DMeasuring7 av range
            rangeDeCititAV.Add(Scan3DMeasuring7AVRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring7LSLRange);
            rangeDeCititLslUsl.Add(Scan3DMeasuring7USLRange);




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

            txt_rangeProfileScanPoint0ValLSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint0ValLSL.ToString();
            txt_rangeProfileScanPoint0ValUSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint0ValUSL.ToString();

            txt_rangeProfileScanPoint1ValLSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint1ValLSL.ToString();
            txt_rangeProfileScanPoint1ValUSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint1ValUSL.ToString();

            txt_rangeProfileScanPoint2ValLSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint2ValLSL.ToString();
            txt_rangeProfileScanPoint2ValUSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint2ValUSL.ToString();

            txt_rangeProfileScanPoint3ValLSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint3ValLSL.ToString();
            txt_rangeProfileScanPoint3ValUSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint3ValUSL.ToString();

            txt_rangeProfileScanPoint4ValLSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint4ValLSL.ToString();
            txt_rangeProfileScanPoint4ValUSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint4ValUSL.ToString();

            txt_rangeProfileScanPoint5ValLSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint5ValLSL.ToString();
            txt_rangeProfileScanPoint5ValUSL.Text = Properties.Settings.Default.txt_rangeProfileScanPoint5ValUSL.ToString();

            txt_rangeJoiningForceValLSL.Text = Properties.Settings.Default.txt_rangeJoiningForceValLSL.ToString();
            txt_rangeJoiningForceValUSL.Text = Properties.Settings.Default.txt_rangeJoiningForceValUSL.ToString();

            txt_rangeLaserWeldingSettingDistanceValLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistanceValLSL.ToString();
            txt_rangeLaserWeldingSettingDistanceValUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingSettingDistanceValUSL.ToString();

            txt_rangeLaserWeldingStartPositionValLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingStartPositionValLSL.ToString();
            txt_rangeLaserWeldingStartPositionValUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingStartPositionValUSL.ToString();

            txt_rangeLaserWeldingShutdownDistanceValLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistanceValLSL.ToString();
            txt_rangeLaserWeldingShutdownDistanceValUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingShutdownDistanceValUSL.ToString();

            txt_rangeLaserWeldingWeldingtimeValLSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtimeValLSL.ToString();
            txt_rangeLaserWeldingWeldingtimeValUSL.Text = Properties.Settings.Default.txt_rangeLaserWeldingWeldingtimeValUSL.ToString();

            txt_rangeForceValLSL.Text = Properties.Settings.Default.txt_rangeForceValLSL.ToString();
            txt_rangeForceValUSL.Text = Properties.Settings.Default.txt_rangeForceValUSL.ToString();

            txt_rangeLeaktestStartPressureValLSL.Text = Properties.Settings.Default.txt_rangeLeaktestStartPressureValLSL.ToString();
            txt_rangeLeaktestStartPressureValUSL.Text = Properties.Settings.Default.txt_rangeLeaktestStartPressureValUSL.ToString();

            txt_rangeLeaktestLeakageValLSL.Text = Properties.Settings.Default.txt_rangeLeaktestLeakageValLSL.ToString();
            txt_rangeLeaktestLeakageValUSL.Text = Properties.Settings.Default.txt_rangeLeaktestLeakageValUSL.ToString();

            txt_rangeEOLM01ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM01ValLSL.ToString();
            txt_rangeEOLM01ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM01ValUSL.ToString();

            txt_rangeEOLM02ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM02ValLSL.ToString();
            txt_rangeEOLM02ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM02ValUSL.ToString();

            txt_rangeEOLM03ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM03ValLSL.ToString();
            txt_rangeEOLM03ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM03ValUSL.ToString();

            txt_rangeEOLM04ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM04ValLSL.ToString();
            txt_rangeEOLM04ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM04ValUSL.ToString();

            txt_rangeEOLM05ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM05ValLSL.ToString();
            txt_rangeEOLM05ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM05ValUSL.ToString();

            txt_rangeEOLM06ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM06ValLSL.ToString();
            txt_rangeEOLM06ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM06ValUSL.ToString();

            txt_rangeEOLM07ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM07ValLSL.ToString();
            txt_rangeEOLM07ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM07ValUSL.ToString();

            txt_rangeEOLM08ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM08ValLSL.ToString();
            txt_rangeEOLM08ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM08ValUSL.ToString();

            txt_rangeEOLM09ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM09ValLSL.ToString();
            txt_rangeEOLM09ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM09ValUSL.ToString();

            txt_rangeEOLM10ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM10ValLSL.ToString();
            txt_rangeEOLM10ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM10ValUSL.ToString();

            txt_rangeEOLM11ValLSL.Text = Properties.Settings.Default.txt_rangeEOLM11ValLSL.ToString();
            txt_rangeEOLM11ValUSL.Text = Properties.Settings.Default.txt_rangeEOLM11ValUSL.ToString();

            txt_rangeScan3DMeasuring0ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring0ValLSL.ToString();
            txt_rangeScan3DMeasuring0ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring0ValUSL.ToString();

            txt_rangeScan3DMeasuring1ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring1ValLSL.ToString();
            txt_rangeScan3DMeasuring1ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring1ValUSL.ToString();

            txt_rangeScan3DMeasuring2ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring2ValLSL.ToString();
            txt_rangeScan3DMeasuring2ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring2ValUSL.ToString();

            txt_rangeScan3DMeasuring3ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring3ValLSL.ToString();
            txt_rangeScan3DMeasuring3ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring3ValUSL.ToString();

            txt_rangeScan3DMeasuring4ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring4ValLSL.ToString();
            txt_rangeScan3DMeasuring4ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring4ValUSL.ToString();

            txt_rangeScan3DMeasuring5ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring5ValLSL.ToString();
            txt_rangeScan3DMeasuring5ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring5ValUSL.ToString();

            txt_rangeScan3DMeasuring6ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring6ValLSL.ToString();
            txt_rangeScan3DMeasuring6ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring6ValUSL.ToString();

            txt_rangeScan3DMeasuring7ValLSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring7ValLSL.ToString();
            txt_rangeScan3DMeasuring7ValUSL.Text = Properties.Settings.Default.txt_rangeScan3DMeasuring7ValUSL.ToString();


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

        public void ComparareValoriProfileScanPoint0LSL()
        {
            for (int i = 0; i < ProfileScanPoint0LSL.Count; i++)
            {
                if (ProfileScanPoint0LSL[i].Equals(txt_rangeProfileScanPoint0ValLSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint0LSL.Text = "OK";
                    lbl_statusProfileScanPoint0LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint0LSL.Text = "NOK";
                    lbl_statusProfileScanPoint0LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint0USL()
        {
            for (int i = 0; i < ProfileScanPoint0USL.Count; i++)
            {
                if (ProfileScanPoint0USL[i].Equals(txt_rangeProfileScanPoint0ValUSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint0USL.Text = "OK";
                    lbl_statusProfileScanPoint0USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint0USL.Text = "NOK";
                    lbl_statusProfileScanPoint0USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriProfileScanPoint1LSL()
        {
            for (int i = 0; i < ProfileScanPoint1LSL.Count; i++)
            {
                if (ProfileScanPoint1LSL[i].Equals(txt_rangeProfileScanPoint1ValLSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint1LSL.Text = "OK";
                    lbl_statusProfileScanPoint1LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint1LSL.Text = "NOK";
                    lbl_statusProfileScanPoint1LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint1USL()
        {
            for (int i = 0; i < ProfileScanPoint1USL.Count; i++)
            {
                if (ProfileScanPoint1USL[i].Equals(txt_rangeProfileScanPoint1ValUSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint1USL.Text = "OK";
                    lbl_statusProfileScanPoint1USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint1USL.Text = "NOK";
                    lbl_statusProfileScanPoint1USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint2LSL()
        {
            for (int i = 0; i < ProfileScanPoint2LSL.Count; i++)
            {
                if (ProfileScanPoint2LSL[i].Equals(txt_rangeProfileScanPoint2ValLSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint2LSL.Text = "OK";
                    lbl_statusProfileScanPoint2LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint2LSL.Text = "NOK";
                    lbl_statusProfileScanPoint2LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint2USL()
        {
            for (int i = 0; i < ProfileScanPoint2USL.Count; i++)
            {
                if (ProfileScanPoint2USL[i].Equals(txt_rangeProfileScanPoint2ValUSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint2USL.Text = "OK";
                    lbl_statusProfileScanPoint2USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint2USL.Text = "NOK";
                    lbl_statusProfileScanPoint2USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint3LSL()
        {
            for (int i = 0; i < ProfileScanPoint3LSL.Count; i++)
            {
                if (ProfileScanPoint3LSL[i].Equals(txt_rangeProfileScanPoint3ValLSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint3LSL.Text = "OK";
                    lbl_statusProfileScanPoint3LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint3LSL.Text = "NOK";
                    lbl_statusProfileScanPoint3LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint3USL()
        {
            for (int i = 0; i < ProfileScanPoint3USL.Count; i++)
            {
                if (ProfileScanPoint3USL[i].Equals(txt_rangeProfileScanPoint3ValUSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint3USL.Text = "OK";
                    lbl_statusProfileScanPoint3USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint3USL.Text = "NOK";
                    lbl_statusProfileScanPoint3USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint4LSL()
        {
            for (int i = 0; i < ProfileScanPoint4LSL.Count; i++)
            {
                if (ProfileScanPoint4LSL[i].Equals(txt_rangeProfileScanPoint4ValLSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint4LSL.Text = "OK";
                    lbl_statusProfileScanPoint4LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint4LSL.Text = "NOK";
                    lbl_statusProfileScanPoint4LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint4USL()
        {
            for (int i = 0; i < ProfileScanPoint4USL.Count; i++)
            {
                if (ProfileScanPoint4USL[i].Equals(txt_rangeProfileScanPoint4ValUSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint4USL.Text = "OK";
                    lbl_statusProfileScanPoint4USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint4USL.Text = "NOK";
                    lbl_statusProfileScanPoint4USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint5LSL()
        {
            for (int i = 0; i < ProfileScanPoint5LSL.Count; i++)
            {
                if (ProfileScanPoint5LSL[i].Equals(txt_rangeProfileScanPoint5ValLSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint5LSL.Text = "OK";
                    lbl_statusProfileScanPoint5LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint5LSL.Text = "NOK";
                    lbl_statusProfileScanPoint5LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriProfileScanPoint5USL()
        {
            for (int i = 0; i < ProfileScanPoint5USL.Count; i++)
            {
                if (ProfileScanPoint5USL[i].Equals(txt_rangeProfileScanPoint5ValUSL.Text.ToString()))
                {
                    lbl_statusProfileScanPoint5USL.Text = "OK";
                    lbl_statusProfileScanPoint5USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusProfileScanPoint5USL.Text = "NOK";
                    lbl_statusProfileScanPoint5USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriJoiningForceLSL()
        {
            for (int i = 0; i < JoiningForceLSL.Count; i++)
            {
                if (JoiningForceLSL[i].Equals(txt_rangeJoiningForceValLSL.Text.ToString()))
                {
                    lbl_statusJoiningForceLSL.Text = "OK";
                    lbl_statusJoiningForceLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusJoiningForceLSL.Text = "NOK";
                    lbl_statusJoiningForceLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriJoiningForceUSL()
        {
            for (int i = 0; i < JoiningForceUSL.Count; i++)
            {
                if (JoiningForceUSL[i].Equals(txt_rangeJoiningForceValUSL.Text.ToString()))
                {
                    lbl_statusJoiningForceUSL.Text = "OK";
                    lbl_statusJoiningForceUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusJoiningForceUSL.Text = "NOK";
                    lbl_statusJoiningForceUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLaserWeldingSettingDistanceLSL()
        {
            for (int i = 0; i < LaserWeldingSettingDistanceLSL.Count; i++)
            {
                if (LaserWeldingSettingDistanceLSL[i].Equals(txt_rangeLaserWeldingSettingDistanceValLSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingSettingDistanceLSL.Text = "OK";
                    lbl_statusLaserWeldingSettingDistanceLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingSettingDistanceLSL.Text = "NOK";
                    lbl_statusLaserWeldingSettingDistanceLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLaserWeldingSettingDistanceUSL()
        {
            for (int i = 0; i < LaserWeldingSettingDistanceUSL.Count; i++)
            {
                if (LaserWeldingSettingDistanceUSL[i].Equals(txt_rangeLaserWeldingSettingDistanceValUSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingSettingDistanceUSL.Text = "OK";
                    lbl_statusLaserWeldingSettingDistanceUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingSettingDistanceUSL.Text = "NOK";
                    lbl_statusLaserWeldingSettingDistanceUSL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriLaserWeldingStartPositionLSL()
        {
            for (int i = 0; i < LaserWeldingStartPositionLSL.Count; i++)
            {
                if (LaserWeldingStartPositionLSL[i].Equals(txt_rangeLaserWeldingStartPositionValLSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingStartPositionLSL.Text = "OK";
                    lbl_statusLaserWeldingStartPositionLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingStartPositionLSL.Text = "NOK";
                    lbl_statusLaserWeldingStartPositionLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLaserWeldingStartPositionUSL()
        {
            for (int i = 0; i < LaserWeldingStartPositionUSL.Count; i++)
            {
                if (LaserWeldingStartPositionUSL[i].Equals(txt_rangeLaserWeldingStartPositionValUSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingStartPositionUSL.Text = "OK";
                    lbl_statusLaserWeldingStartPositionUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingStartPositionUSL.Text = "NOK";
                    lbl_statusLaserWeldingStartPositionUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLaserWeldingShutdownDistanceLSL()
        {
            for (int i = 0; i < LaserWeldingShutdownDistanceLSL.Count; i++)
            {
                if (LaserWeldingShutdownDistanceLSL[i].Equals(txt_rangeLaserWeldingShutdownDistanceValLSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingShutdownDistanceLSL.Text = "OK";
                    lbl_statusLaserWeldingShutdownDistanceLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingShutdownDistanceLSL.Text = "NOK";
                    lbl_statusLaserWeldingShutdownDistanceLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLaserWeldingShutdownDistanceUSL()
        {
            for (int i = 0; i < LaserWeldingShutdownDistanceUSL.Count; i++)
            {
                if (LaserWeldingShutdownDistanceUSL[i].Equals(txt_rangeLaserWeldingShutdownDistanceValUSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingShutdownDistanceUSL.Text = "OK";
                    lbl_statusLaserWeldingShutdownDistanceUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingShutdownDistanceUSL.Text = "NOK";
                    lbl_statusLaserWeldingShutdownDistanceUSL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriLaserWeldingWeldingtimeLSL()
        {
            for (int i = 0; i < LaserWeldingWeldingtimeLSL.Count; i++)
            {
                if (LaserWeldingWeldingtimeLSL[i].Equals(txt_rangeLaserWeldingWeldingtimeValLSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingWeldingtimeLSL.Text = "OK";
                    lbl_statusLaserWeldingWeldingtimeLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingWeldingtimeLSL.Text = "NOK";
                    lbl_statusLaserWeldingWeldingtimeLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLaserWeldingWeldingtimeUSL()
        {
            for (int i = 0; i < LaserWeldingWeldingtimeUSL.Count; i++)
            {
                if (LaserWeldingWeldingtimeUSL[i].Equals(txt_rangeLaserWeldingWeldingtimeValUSL.Text.ToString()))
                {
                    lbl_statusLaserWeldingWeldingtimeUSL.Text = "OK";
                    lbl_statusLaserWeldingWeldingtimeUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLaserWeldingWeldingtimeUSL.Text = "NOK";
                    lbl_statusLaserWeldingWeldingtimeUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriForceLSL()
        {
            for (int i = 0; i < ForceLSL.Count; i++)
            {
                if (ForceLSL[i].Equals(txt_rangeForceValLSL.Text.ToString()))
                {
                    lbl_statusForceLSL.Text = "OK";
                    lbl_statusForceLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusForceLSL.Text = "NOK";
                    lbl_statusForceLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriForceUSL()
        {
            for (int i = 0; i < ForceUSL.Count; i++)
            {
                if (ForceUSL[i].Equals(txt_rangeForceValUSL.Text.ToString()))
                {
                    lbl_statusForceUSL.Text = "OK";
                    lbl_statusForceUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusForceUSL.Text = "NOK";
                    lbl_statusForceUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLeaktestStartPressureLSL()
        {
            for (int i = 0; i < LeaktestStartPressureLSL.Count; i++)
            {
                if (LeaktestStartPressureLSL[i].Equals(txt_rangeLeaktestStartPressureValLSL.Text.ToString()))
                {
                    lbl_statusLeaktestStartPressureLSL.Text = "OK";
                    lbl_statusLeaktestStartPressureLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLeaktestStartPressureLSL.Text = "NOK";
                    lbl_statusLeaktestStartPressureLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLeaktestStartPressureUSL()
        {
            for (int i = 0; i < LeaktestStartPressureUSL.Count; i++)
            {
                if (LeaktestStartPressureUSL[i].Equals(txt_rangeLeaktestStartPressureValUSL.Text.ToString()))
                {
                    lbl_statusLeaktestStartPressureUSL.Text = "OK";
                    lbl_statusLeaktestStartPressureUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLeaktestStartPressureUSL.Text = "NOK";
                    lbl_statusLeaktestStartPressureUSL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriLeaktestLeakageLSL()
        {
            for (int i = 0; i < LeaktestLeakageLSL.Count; i++)
            {
                if (LeaktestLeakageLSL[i].Equals(txt_rangeLeaktestLeakageValLSL.Text.ToString()))
                {
                    lbl_statusLeaktestLeakageLSL.Text = "OK";
                    lbl_statusLeaktestLeakageLSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLeaktestLeakageLSL.Text = "NOK";
                    lbl_statusLeaktestLeakageLSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriLeaktestLeakageUSL()
        {
            for (int i = 0; i < LeaktestLeakageUSL.Count; i++)
            {
                if (LeaktestLeakageUSL[i].Equals(txt_rangeLeaktestLeakageValUSL.Text.ToString()))
                {
                    lbl_statusLeaktestLeakageUSL.Text = "OK";
                    lbl_statusLeaktestLeakageUSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusLeaktestLeakageUSL.Text = "NOK";
                    lbl_statusLeaktestLeakageUSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM01LSL()
        {
            for (int i = 0; i < EOLM01LSL.Count; i++)
            {
                if (EOLM01LSL[i].Equals(txt_rangeEOLM01ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM01LSL.Text = "OK";
                    lbl_statusEOLM01LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM01LSL.Text = "NOK";
                    lbl_statusEOLM01LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM01USL()
        {
            for (int i = 0; i < EOLM01USL.Count; i++)
            {
                if (EOLM01USL[i].Equals(txt_rangeEOLM01ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM01USL.Text = "OK";
                    lbl_statusEOLM01USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM01USL.Text = "NOK";
                    lbl_statusEOLM01USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriEOLM02LSL()
        {
            for (int i = 0; i < EOLM02LSL.Count; i++)
            {
                if (EOLM02LSL[i].Equals(txt_rangeEOLM02ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM02LSL.Text = "OK";
                    lbl_statusEOLM02LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM02LSL.Text = "NOK";
                    lbl_statusEOLM02LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM02USL()
        {
            for (int i = 0; i < EOLM02USL.Count; i++)
            {
                if (EOLM02USL[i].Equals(txt_rangeEOLM02ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM02USL.Text = "OK";
                    lbl_statusEOLM02USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM02USL.Text = "NOK";
                    lbl_statusEOLM02USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM03LSL()
        {
            for (int i = 0; i < EOLM03LSL.Count; i++)
            {
                if (EOLM03LSL[i].Equals(txt_rangeEOLM03ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM03LSL.Text = "OK";
                    lbl_statusEOLM03LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM03LSL.Text = "NOK";
                    lbl_statusEOLM03LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM03USL()
        {
            for (int i = 0; i < EOLM03USL.Count; i++)
            {
                if (EOLM03USL[i].Equals(txt_rangeEOLM03ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM03USL.Text = "OK";
                    lbl_statusEOLM03USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM03USL.Text = "NOK";
                    lbl_statusEOLM03USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM04LSL()
        {
            for (int i = 0; i < EOLM04LSL.Count; i++)
            {
                if (EOLM04LSL[i].Equals(txt_rangeEOLM04ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM04LSL.Text = "OK";
                    lbl_statusEOLM04LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM04LSL.Text = "NOK";
                    lbl_statusEOLM04LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM04USL()
        {
            for (int i = 0; i < EOLM04USL.Count; i++)
            {
                if (EOLM04USL[i].Equals(txt_rangeEOLM04ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM04USL.Text = "OK";
                    lbl_statusEOLM04USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM04USL.Text = "NOK";
                    lbl_statusEOLM04USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM05LSL()
        {
            for (int i = 0; i < EOLM05LSL.Count; i++)
            {
                if (EOLM05LSL[i].Equals(txt_rangeEOLM05ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM05LSL.Text = "OK";
                    lbl_statusEOLM05LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM05LSL.Text = "NOK";
                    lbl_statusEOLM05LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM05USL()
        {
            for (int i = 0; i < EOLM05USL.Count; i++)
            {
                if (EOLM05USL[i].Equals(txt_rangeEOLM05ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM05USL.Text = "OK";
                    lbl_statusEOLM05USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM05USL.Text = "NOK";
                    lbl_statusEOLM05USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM06LSL()
        {
            for (int i = 0; i < EOLM06LSL.Count; i++)
            {
                if (EOLM06LSL[i].Equals(txt_rangeEOLM06ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM06LSL.Text = "OK";
                    lbl_statusEOLM06LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM06LSL.Text = "NOK";
                    lbl_statusEOLM06LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM06USL()
        {
            for (int i = 0; i < EOLM06USL.Count; i++)
            {
                if (EOLM06USL[i].Equals(txt_rangeEOLM06ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM06USL.Text = "OK";
                    lbl_statusEOLM06USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM06USL.Text = "NOK";
                    lbl_statusEOLM06USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM07LSL()
        {
            for (int i = 0; i < EOLM07LSL.Count; i++)
            {
                if (EOLM07LSL[i].Equals(txt_rangeEOLM07ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM07LSL.Text = "OK";
                    lbl_statusEOLM07LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM07LSL.Text = "NOK";
                    lbl_statusEOLM07LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM07USL()
        {
            for (int i = 0; i < EOLM07USL.Count; i++)
            {
                if (EOLM07USL[i].Equals(txt_rangeEOLM07ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM07USL.Text = "OK";
                    lbl_statusEOLM07USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM07USL.Text = "NOK";
                    lbl_statusEOLM07USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM08LSL()
        {
            for (int i = 0; i < EOLM08LSL.Count; i++)
            {
                if (EOLM08LSL[i].Equals(txt_rangeEOLM08ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM08LSL.Text = "OK";
                    lbl_statusEOLM08LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM08LSL.Text = "NOK";
                    lbl_statusEOLM08LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM08USL()
        {
            for (int i = 0; i < EOLM08USL.Count; i++)
            {
                if (EOLM08USL[i].Equals(txt_rangeEOLM08ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM08USL.Text = "OK";
                    lbl_statusEOLM08USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM08USL.Text = "NOK";
                    lbl_statusEOLM08USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM09LSL()
        {
            for (int i = 0; i < EOLM09LSL.Count; i++)
            {
                if (EOLM09LSL[i].Equals(txt_rangeEOLM09ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM09LSL.Text = "OK";
                    lbl_statusEOLM09LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM09LSL.Text = "NOK";
                    lbl_statusEOLM09LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM09USL()
        {
            for (int i = 0; i < EOLM09USL.Count; i++)
            {
                if (EOLM09USL[i].Equals(txt_rangeEOLM09ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM09USL.Text = "OK";
                    lbl_statusEOLM09USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM09USL.Text = "NOK";
                    lbl_statusEOLM09USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM10LSL()
        {
            for (int i = 0; i < EOLM10LSL.Count; i++)
            {
                if (EOLM10LSL[i].Equals(txt_rangeEOLM10ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM10LSL.Text = "OK";
                    lbl_statusEOLM10LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM10LSL.Text = "NOK";
                    lbl_statusEOLM10LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM10USL()
        {
            for (int i = 0; i < EOLM10USL.Count; i++)
            {
                if (EOLM10USL[i].Equals(txt_rangeEOLM10ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM10USL.Text = "OK";
                    lbl_statusEOLM10USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM10USL.Text = "NOK";
                    lbl_statusEOLM10USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM11LSL()
        {
            for (int i = 0; i < EOLM11LSL.Count; i++)
            {
                if (EOLM11LSL[i].Equals(txt_rangeEOLM11ValLSL.Text.ToString()))
                {
                    lbl_statusEOLM11LSL.Text = "OK";
                    lbl_statusEOLM11LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM11LSL.Text = "NOK";
                    lbl_statusEOLM11LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriEOLM11USL()
        {
            for (int i = 0; i < EOLM11USL.Count; i++)
            {
                if (EOLM11USL[i].Equals(txt_rangeEOLM11ValUSL.Text.ToString()))
                {
                    lbl_statusEOLM11USL.Text = "OK";
                    lbl_statusEOLM11USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusEOLM11USL.Text = "NOK";
                    lbl_statusEOLM11USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring0LSL()
        {
            for (int i = 0; i < Scan3DMeasuring0LSL.Count; i++)
            {
                if (Scan3DMeasuring0LSL[i].Equals(txt_rangeScan3DMeasuring0ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring0LSL.Text = "OK";
                    lbl_statusScan3DMeasuring0LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring0LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring0LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring0USL()
        {
            for (int i = 0; i < Scan3DMeasuring0USL.Count; i++)
            {
                if (Scan3DMeasuring0USL[i].Equals(txt_rangeScan3DMeasuring0ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring0USL.Text = "OK";
                    lbl_statusScan3DMeasuring0USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring0USL.Text = "NOK";
                    lbl_statusScan3DMeasuring0USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriScan3DMeasuring1LSL()
        {
            for (int i = 0; i < Scan3DMeasuring1LSL.Count; i++)
            {
                if (Scan3DMeasuring1LSL[i].Equals(txt_rangeScan3DMeasuring1ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring1LSL.Text = "OK";
                    lbl_statusScan3DMeasuring1LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring1LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring1LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring1USL()
        {
            for (int i = 0; i < Scan3DMeasuring1USL.Count; i++)
            {
                if (Scan3DMeasuring1USL[i].Equals(txt_rangeScan3DMeasuring1ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring1USL.Text = "OK";
                    lbl_statusScan3DMeasuring1USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring1USL.Text = "NOK";
                    lbl_statusScan3DMeasuring1USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring2LSL()
        {
            for (int i = 0; i < Scan3DMeasuring2LSL.Count; i++)
            {
                if (Scan3DMeasuring2LSL[i].Equals(txt_rangeScan3DMeasuring2ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring2LSL.Text = "OK";
                    lbl_statusScan3DMeasuring2LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring2LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring2LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring2USL()
        {
            for (int i = 0; i < Scan3DMeasuring2USL.Count; i++)
            {
                if (Scan3DMeasuring2USL[i].Equals(txt_rangeScan3DMeasuring2ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring2USL.Text = "OK";
                    lbl_statusScan3DMeasuring2USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring2USL.Text = "NOK";
                    lbl_statusScan3DMeasuring2USL.BackColor = Color.Red;
                    break;
                }
            }

        }



        public void ComparareValoriScan3DMeasuring3LSL()
        {
            for (int i = 0; i < Scan3DMeasuring3LSL.Count; i++)
            {
                if (Scan3DMeasuring3LSL[i].Equals(txt_rangeScan3DMeasuring3ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring3LSL.Text = "OK";
                    lbl_statusScan3DMeasuring3LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring3LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring3LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring3USL()
        {
            for (int i = 0; i < Scan3DMeasuring3USL.Count; i++)
            {
                if (Scan3DMeasuring3USL[i].Equals(txt_rangeScan3DMeasuring3ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring3USL.Text = "OK";
                    lbl_statusScan3DMeasuring3USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring3USL.Text = "NOK";
                    lbl_statusScan3DMeasuring3USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring4LSL()
        {
            for (int i = 0; i < Scan3DMeasuring4LSL.Count; i++)
            {
                if (Scan3DMeasuring4LSL[i].Equals(txt_rangeScan3DMeasuring4ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring4LSL.Text = "OK";
                    lbl_statusScan3DMeasuring4LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring4LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring4LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring4USL()
        {
            for (int i = 0; i < Scan3DMeasuring4USL.Count; i++)
            {
                if (Scan3DMeasuring4USL[i].Equals(txt_rangeScan3DMeasuring4ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring4USL.Text = "OK";
                    lbl_statusScan3DMeasuring4USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring4USL.Text = "NOK";
                    lbl_statusScan3DMeasuring4USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring5LSL()
        {
            for (int i = 0; i < Scan3DMeasuring5LSL.Count; i++)
            {
                if (Scan3DMeasuring5LSL[i].Equals(txt_rangeScan3DMeasuring5ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring5LSL.Text = "OK";
                    lbl_statusScan3DMeasuring5LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring5LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring5LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring5USL()
        {
            for (int i = 0; i < Scan3DMeasuring5USL.Count; i++)
            {
                if (Scan3DMeasuring5USL[i].Equals(txt_rangeScan3DMeasuring5ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring5USL.Text = "OK";
                    lbl_statusScan3DMeasuring5USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring5USL.Text = "NOK";
                    lbl_statusScan3DMeasuring5USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring6LSL()
        {
            for (int i = 0; i < Scan3DMeasuring6LSL.Count; i++)
            {
                if (Scan3DMeasuring6LSL[i].Equals(txt_rangeScan3DMeasuring6ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring6LSL.Text = "OK";
                    lbl_statusScan3DMeasuring6LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring6LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring6LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring6USL()
        {
            for (int i = 0; i < Scan3DMeasuring6USL.Count; i++)
            {
                if (Scan3DMeasuring6USL[i].Equals(txt_rangeScan3DMeasuring6ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring6USL.Text = "OK";
                    lbl_statusScan3DMeasuring6USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring6USL.Text = "NOK";
                    lbl_statusScan3DMeasuring6USL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring7LSL()
        {
            for (int i = 0; i < Scan3DMeasuring7LSL.Count; i++)
            {
                if (Scan3DMeasuring7LSL[i].Equals(txt_rangeScan3DMeasuring7ValLSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring7LSL.Text = "OK";
                    lbl_statusScan3DMeasuring7LSL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring7LSL.Text = "NOK";
                    lbl_statusScan3DMeasuring7LSL.BackColor = Color.Red;
                    break;
                }
            }

        }

        public void ComparareValoriScan3DMeasuring7USL()
        {
            for (int i = 0; i < Scan3DMeasuring7USL.Count; i++)
            {
                if (Scan3DMeasuring7USL[i].Equals(txt_rangeScan3DMeasuring7ValUSL.Text.ToString()))
                {
                    lbl_statusScan3DMeasuring7USL.Text = "OK";
                    lbl_statusScan3DMeasuring7USL.BackColor = Color.GreenYellow;
                }


                else
                {
                    lbl_statusScan3DMeasuring7USL.Text = "NOK";
                    lbl_statusScan3DMeasuring7USL.BackColor = Color.Red;
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


