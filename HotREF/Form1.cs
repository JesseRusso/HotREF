//Created by Jesse Russo 2019
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
        string excelFilePath;
        string zone = "7A";
        private string proposedAddress;
        private string directoryString;
        public string excelPath { get; private set; }
         
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

        private void Button1_Click_1(object sender, EventArgs e)
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
                excelPath = excelFilePath;
                worksheetTextBox.Text = excelAddress;
                proposedAddress = excelAddress;
            }
            ofd.Dispose();
        }

        private void CreateProp_Click(object sender, EventArgs e)
        {
            CreateProp cp = new CreateProp(excelFilePath, template);

            cp.FindID(template);
            cp.ChangeEquipment();
            cp.ChangeSpecs();
            cp.ChangeAddress(proposedAddress);
            cp.ChangeWalls();
            cp.CheckCeilings();
            cp.ChangeFloors();
            cp.ExtraFloors();
            cp.ExtraCeilings();
            cp.CheckVaults();
            cp.ExtraWalls();
            cp.ChangeBasment();
            cp.GasDHW();
            cp.ElectricDHW();

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

