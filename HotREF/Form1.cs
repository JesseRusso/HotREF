﻿//Created by Jesse Russo 2019
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HotREF
{
    public partial class Form1 : Form
    {

        XDocument propHouse;
        
        string filePath;
        string tankEF = "0.58";
        string furnaceOutput;
        string fanPower = "125.3";
        string ventilation = "57.2";
        string windowSize;
        string doorSize;
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
                filePath = ofd.FileName;
                textBox1.Text = filePath;
                propHouse = XDocument.Load(filePath);
            }
            ofd.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

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
                CreateRef cr = new CreateRef(propHouse);
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
            ofd.Title = "Select workbook";
            ofd.Filter = "*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                var sheet = ofd.FileName;
            }
            ofd.Dispose();
        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
