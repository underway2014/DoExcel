using DoExcel.Helper;
using DoExcel.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DoExcel
{
    public partial class Form1 : Form
    {

        public string root = "d:/DoExcel/";
        /// <summary>
        /// 数据源
        /// </summary>
        public string originSourcePath = "data.xls";

        /// <summary>
        /// 生成模板地址
        /// </summary>
        public static string TemplatePath = "model.xls";

        private object excelHelper;


        public Form1()
        {
            InitializeComponent();

            verLabel.Text = "V:" + getToolVersion();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            labelStatus.Text = "处理中...";
            ExcelHelper excelHelper = new ExcelHelper();

            List<Person> list = excelHelper.ReadFromExcelFile(root + originSourcePath);
            string timeStr = DateTime.Now.ToString("yyyyMMddHHmmss");
            string rootTarget = root + timeStr + "/";
            if (!Directory.Exists(rootTarget))
            {
                Directory.CreateDirectory(rootTarget);
            }

            foreach (Person person in list)
            {
                string target = rootTarget + person.Name + ".xls";
                excelHelper.CopyExcel(root + TemplatePath, target, person);
            }

            labelStatus.Text = "处理结束，共处理" + list.Count + "条数据！";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelHelper excelHelper = new ExcelHelper();

            excelHelper.WriteToExcel("D:/c#project/vbProject/excel/demo1.xls");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelHelper excelHelper = new ExcelHelper();
            Person person = new Person();
            //excelHelper.CopyExcel("D:/c#project/vbProject/excel/copy1.xls", person);
        }

        public string getToolVersion()
        {
            string major = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString();
            string minor = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString();
            string build = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString();
            string revision = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Revision.ToString();

            return major + "." + minor + "." + build + "." + revision;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
