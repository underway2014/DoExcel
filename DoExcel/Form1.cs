﻿using DoExcel.Helper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DoExcel
{
    public partial class Form1 : Form
    {
        public string originSourcePath = "D:/c#project/vbProject/excel/libin.xls";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelHelper excelHelper = new ExcelHelper();

            excelHelper.ReadFromExcelFile(originSourcePath);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelHelper excelHelper = new ExcelHelper();

            excelHelper.WriteToExcel("D:/c#project/vbProject/excel/write.xls");
        }
    }
}