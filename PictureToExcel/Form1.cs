using System;
using System.IO;
using System.Threading;
using System.Drawing;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PictureToExcel
{
    public partial class Form1 : Form
    {
        private string picPath;
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            progressBar1.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(pictureBox1.Image == null)
            {
                MessageBox.Show("先选择图片", "提示");
                return;
            }

            pictureBox1.Enabled = false;
            button1.Enabled = false;
            progressBar1.Show();
            Thread t1 = new Thread(new ThreadStart(creatMethod));
            t1.IsBackground = true;
            t1.Start();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            pictureBox1_Click(sender, e);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "请选择图片";
            openFileDialog1.Filter = "图片文件(*.jpg,*.jpeg,*.png,*.gif,*.bmp)|*.jpg|*.jpeg|*.png|*.gif|*.bmp";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                picPath = openFileDialog1.FileName.ToString();
            }

            try
            {
                pictureBox1.Image = Image.FromFile(picPath);
                label1.Hide();
            }
            catch
            {
                MessageBox.Show("选择图片加载失败", "警告");
                label1.Show();
            }  
        }

        private void creatMethod()
        {
            //create  
            var file = @"result.xlsx";
            if (File.Exists(file))
            {
                File.Delete(file);
            }
            
            using (var excel = new ExcelPackage(new FileInfo(file)))
            {
                var ws = excel.Workbook.Worksheets.Add("Sheet1");

                using (Bitmap bmp = new Bitmap(picPath))
                {
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = bmp.Height * bmp.Width;
                    progressBar1.Value = 0;

                    for (var y = 0; y < bmp.Height; y++)
                    {
                        ws.Row(y + 1).Height = 5.25;

                        for (var x = 0; x < bmp.Width; x++)
                        {
                            Color color = bmp.GetPixel(x, y);

                            ws.Column(x + 1).Width = 1.00;
                            ws.Cells[y + 1, x + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[y + 1, x + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(color.A, color.R, color.G, color.B));

                            progressBar1.Value++;
                        }
                    }

                }
                excel.Save();
            }

            progressBar1.Hide();
            pictureBox1.Enabled = true;
            button1.Enabled = true;
            MessageBox.Show("转换完成", "提示");
        }
    }
}
