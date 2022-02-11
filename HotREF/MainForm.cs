﻿//Created by Jesse Russo 2019
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Xml.Linq;
using System.Text;
using System.Windows.Forms;
using HotREF.Properties;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace HotREF
{
    public partial class MainForm : Form
    {
        XDocument propHouse;
        private XDocument newHouse;
        string templatePath;
        string excelFilePath;
        string zone = "7A";
        private string proposedAddress;
        private string directoryString;
        public string ExcelPath { get; private set; }
         
        public MainForm()
        {
            InitializeComponent();
        }

        private void SelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select Proposed File";
            ofd.Filter = "House Files (*.h2k)|*.h2k";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                propHouse = XDocument.Load(ofd.FileName);
                directoryString = Path.GetDirectoryName(ofd.FileName);
                string[] pathStrings = Path.GetFileName(ofd.FileName).Split('-');

                if(pathStrings.Length > 2)
                {
                    proposedAddress = pathStrings[0];
                    for(int i = 1; i < pathStrings.Length-1; i++)
                    {
                        proposedAddress += $"-{pathStrings[i]}";
                    }
                }
                else
                {
                    proposedAddress = pathStrings[0];
                }
                textBox1.Text = proposedAddress;
            }
            ofd.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Location = Settings.Default.WindowLocation;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            CreateRef cr = new CreateRef(propHouse,zone,excelFilePath);
            cr.FindID(propHouse);
            propHouse = cr.Remover(propHouse);
            propHouse = cr.AddCode(propHouse);
            propHouse = cr.RChanger(propHouse);
            propHouse = cr.HvacChanger(propHouse);
            propHouse = cr.AddFan(propHouse);
            propHouse = cr.Doors(propHouse);
            propHouse = cr.Windows(propHouse);
            propHouse = cr.HotWater(propHouse);

            MessageBox.Show("Please save and check results", "REF changes made");

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "House File|*.h2k";
            sfd.DefaultExt = "h2k";
            sfd.InitialDirectory = directoryString;
            sfd.FileName = $"{proposedAddress}-REFERENCE";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                propHouse.Save(sfd.FileName);
            }
            sfd.Dispose();
        }

        private void CreateProp_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            XDocument template = new XDocument(XDocument.Load(templatePath));
            CreateProp cp = new CreateProp(excelFilePath, template);
            cp.FindID(template);
            cp.ChangeAddress(proposedAddress);
            try
            {
                cp.ChangeEquipment();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("There was an error. Check furnace and HRV values in excel.",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ChangeSpecs();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("There was an error. Check for typos in intersections/corners and volume/highest ceiling.", 
                    "Something went wrong", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ChangeWalls();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("There was an error. Check for typos in above grade walls and " +
                    "check that the H2K template has all required elements.",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.CheckCeilings();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("There was an error retrieving ceiling data from the template. " +
                    "Check that the H2K template has the required ceilings.",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ChangeFloors();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Error either retrieving floor R values from template, or adding garage " +
                    "floor. Template should have 2 exposed floors with 'cant' and 'garage' in their names",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ExtraFloors();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Unexpected value while adding floors. Have a typo in the EXPOSED FLOORS section?",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ExtraCeilings();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Unexpected value while adding ceilings. Have a typo in the FLAT CEILINGS section?",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.CheckVaults();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Unexpected value while adding vaults. Have a typo in the VAULTS section?",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ExtraWalls();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Unexpected value while adding walls. Have a typo in the ABOVE GRADE WALLS section?",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ChangeBasment();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Error while changing basement. " +
                    "Check template has required basement elements then check spreadsheet for typos.",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.GasDHW();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Error while changing GAS DHW " +
                    "Check for typos in GAS DHW section.",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                cp.ElectricDHW();
            }
            catch
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Error while changing ELECTRIC DHW " +
                    "Check for typos in ELECTRIC DHW section.",
                    "Something went wrong",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            newHouse = cp.GetHouse();

            Cursor = Cursors.Default;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save Generated Proposed House";
            sfd.Filter = " H2K files (*.h2k)| *.h2k";
            sfd.InitialDirectory = Path.GetDirectoryName(excelFilePath);
            sfd.FileName = $"{proposedAddress}-PROPOSED";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                newHouse.Save(sfd.FileName);
            }
            template = null;
            GC.Collect();
        }

        private void TemplateButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select HOT2000 builder template";
            ofd.Filter = "House Files(*.h2k) | *.h2k";
            ofd.InitialDirectory = Settings.Default.TemplateDir;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //template = XDocument.Load(ofd.FileName);
                templatePath = ofd.FileName;
            }
            ofd.Dispose();
        }

        private void ZoneSelectBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            zone = ZoneSelectBox.Text;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Settings.Default.WindowLocation = this.Location;
            Settings.Default.Save();
        }

        private void TemplateDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Choose default directory for H2K templates";
            if(fbd.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.TemplateDir = fbd.SelectedPath;
                Settings.Default.Save();
            }
        }

        private void Choose_Worksheet_Button(object sender, EventArgs e)
        {
            string excelAddress;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select worksheet";
            ofd.Filter = "Excel Files (*.xlsm) | *.xlsm";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                worksheetTextBox.Clear();
                excelFilePath = ofd.FileName.ToString();
                string[] pathStrings = Path.GetFileName(ofd.FileName).Split('-');

                if (pathStrings.Length > 2)
                {
                    excelAddress = pathStrings[0];
                    for (int i = 1; i < pathStrings.Length - 1; i++)
                    {
                        excelAddress += $"-{pathStrings[i]}";
                    }
                }
                else
                {
                    excelAddress = pathStrings[0];
                }
                ExcelPath = excelFilePath;
                worksheetTextBox.Text = excelAddress;
                proposedAddress = excelAddress;
            }
            ofd.Dispose();
        }
    }
}
