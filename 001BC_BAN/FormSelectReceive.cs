﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace BC_BANSONG
{
    public partial class FormSelectReceive : Form
    {
        private int iSelect = 0;

        public FormSelectReceive()
        {
            InitializeComponent();
        }

        public int GetSelected
        {
            get
            {
                return iSelect;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ButtonClick(1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ButtonClick(2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ButtonClick(3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ButtonClick(4);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ButtonClick(0);
        }

        private void ButtonClick(int value)
        {
            iSelect = value;
            this.DialogResult = DialogResult.OK;
        }
    }
}