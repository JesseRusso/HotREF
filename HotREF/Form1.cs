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

namespace HotREF
{
    public partial class Form1 : Form
    {
        XDocument propHouse;
        XDocument template;

        string tankEF = "0.58";
        string furnaceOutput;
        string fanPower = "125.3";
        string ventilation = "57.2";
        string windowSize;
        string doorSize;
        string excelFilePath;
        string zone = "7A";
        private string proposedAddress;
        private string directoryString;
         

        public Form1()
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
                for (int i = 0; i < pathStrings.Length; i++ ) 
                {
                    Debug.WriteLine(pathStrings[i]);
                }
                proposedAddress = pathStrings[0];
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
            if (furnaceOutput == null)
            {
                MessageBox.Show("Please enter a value for furnace output", "Furnace output empty");
            }
            else if (windowSize == null)
            {
                MessageBox.Show("", "Window size required");
            }
            else if (tankEF == null)
            {
                MessageBox.Show("Please enter a value for tank EF", "Tank EF empty");
            }
            else if (fanPower == null)
            {
                MessageBox.Show("Please enter a value for fan power", "Fan power empty");
            }
            else if (propHouse == null)
            {
                MessageBox.Show("House File must be selected", "No house file");
            }
            else if (ventilation == null)
            {
                MessageBox.Show("Please enter a value for ventilation", "Ventilation empty");
            }
            else
            {
                CreateRef cr = new CreateRef(propHouse,zone);
                cr.FindID(propHouse);
                propHouse = cr.Remover(propHouse);
                propHouse = cr.AddCode(propHouse);
                propHouse = cr.RChanger(propHouse);
                propHouse = cr.HvacChanger(propHouse, furnaceOutput);
                propHouse = cr.AddHrv(propHouse, ventilation, fanPower);
                propHouse = cr.Doors(propHouse, doorSize);
                propHouse = cr.Windows(propHouse, windowSize);
                propHouse = cr.HotWater(propHouse, tankEF);

                MessageBox.Show("Please save and check results", "REF changes made");
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
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

        private void TextBox8_TextChanged(object sender, EventArgs e)
        {
            tankEF = textBox8.Text;
        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            windowSize = textBox3.Text;
        }

        private void TextBox4_TextChanged(object sender, EventArgs e)
        {
            furnaceOutput = textBox4.Text;
        }

        private void TextBox7_TextChanged(object sender, EventArgs e)
        {
            fanPower = textBox7.Text;
        }

        private void TextBox6_TextChanged(object sender, EventArgs e)
        {
            ventilation = textBox6.Text;
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            doorSize = textBox2.Text;
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select worksheet";
            ofd.Filter = "Excel Files (*.xlsm) | *.xlsm";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = ofd.FileName.ToString();
                string[] pathStrings = Path.GetFileName(ofd.FileName).Split('-');
                proposedAddress = pathStrings[0];
            }
            ofd.Dispose();
        }


        private void CreateProp_Click(object sender, EventArgs e)
        {
            CreateProp cp = new CreateProp(excelFilePath, template);

            cp.FindID(template);
            cp.ChangeEquipment();
            cp.ChangeSpecs();
            cp.ChangeWalls();
            cp.CheckCeilings();
            cp.ChangeFloors();
            cp.ExtraFloors();
            cp.ExtraCeilings();
            cp.CheckVaults();
            cp.ExtraWalls();
            cp.ChangeBasment();

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save Generated Proposed House";
            sfd.Filter = " H2K files (*.h2k)| *.h2k";
            sfd.InitialDirectory = Path.GetDirectoryName(excelFilePath);
            sfd.FileName = $"{proposedAddress}-PROPOSED";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                template.Save(sfd.FileName);
            }
            template = null;
        }

        private void TemplateButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select HOT2000 builder template";
            ofd.Filter = "House Files(*.h2k) | *.h2k";
            ofd.InitialDirectory = Settings.Default.TemplateDir;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                template = XDocument.Load(ofd.FileName);
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
            fbd.Description = "Choose default template directory";
            if(fbd.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.TemplateDir = fbd.SelectedPath;
                Settings.Default.Save();
            }
        }
    }
}

