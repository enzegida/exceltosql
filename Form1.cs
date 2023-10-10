using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToSql
{
    public partial class SciptMaker : Form
    {
        public SciptMaker()
        {
            InitializeComponent();
        }

        private void SciptMaker_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try{
                FileInfo existingFile = new FileInfo(textBox1.Text);

                string directoryPath = System.IO.Path.GetDirectoryName(textBox1.Text);
                MessageBox.Show("Diretctory is :" + directoryPath.ToString());


                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    //int colCount = worksheet.Dimension.End.Column;
                    int rowCount = worksheet.Dimension.End.Row;
                    int colCount = worksheet.Dimension.End.Column;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        //create a bool
                        bool RowIsEmpty = true;

                        for (int col = 1; col <= colCount; col++)
                        {
                            //check if the cell is empty or not
                            if (worksheet.Cells[row, col].Value != null)
                            {
                                RowIsEmpty = false;
                            }
                        }

                        //display result
                        if (RowIsEmpty)
                        {
                            MessageBox.Show("Row " + row + " is empty.<br>");
                        }
                    }


                    //MessageBox.Show(rowCount.ToString());
                    //MessageBox.Show(comboBox1.Text);
                    //if(comboBox1.Text == "CVTS")
                    //{
                    StringBuilder sb = new StringBuilder();
                    for (int row = 1; row < rowCount + 1; row++)
                    {

                        sb.AppendLine("INSERT INTO [dbo].[Transactions]\r\n" +
                            "([batch_code]\r\n" +
                            ",[tracking_number]\r\n" +
                            "           ,[card_number]\r\n           " +
                            ",[other_card_number]\r\n           " +
                            ",[cardholder_name]\r\n           " +
                            ",[other_cardholder_name]\r\n           " +
                            ",[address1]\r\n           " +
                            ",[address2]\r\n           " +
                            ",[address3]\r\n           " +
                            ",[address4]\r\n           " +
                            ",[zip_code]\r\n           " +
                            ",[phone1]\r\n           " +
                            ",[phone2]\r\n           " +
                            ",[attempt_counter]\r\n           " +
                            ",[area_code]\r\n           " +
                            ",[emboss_request_code]\r\n           " +
                            ",[envelope_code]\r\n           " +
                            ",[delivery_attempt_code]\r\n           " +
                            ",[plastic_code]\r\n           " +
                            ",[voucher_amount]\r\n           " +
                            ",[received_by]\r\n           " +
                            ",[relationship]\r\n           " +
                            ",[user_field1]\r\n           " +
                            ",[user_field2]\r\n           " +
                            ",[due_date]\r\n           " +
                            ",[current_status]\r\n           " +
                            ",[status_date]\r\n           " +
                            ",[courier_code]\r\n           " +
                            ",[sq_due_date]\r\n           " +
                            ",[source]\r\n           " +
                            ",[cust_num]\r\n           " +
                            ",[subtrack_flag]\r\n           " +
                            ",[subtrack_date]\r\n           " +
                            ",[Expiry_date]\r\n           " +
                            ",[courier_tracking_number]\r\n           " +
                            ",[seq])\r\n     " +
                            "VALUES\r\n           " +
                            "('" + worksheet.Cells[row, 1].Value + "'\r\n           " +
                            ",'" + worksheet.Cells[row, 1].Value + "'\r\n           " +
                            ",'" + worksheet.Cells[row, 2].Value + "'\r\n           " +
                            ",''\r\n           " +
                            ",'" + worksheet.Cells[row, 3].Value + "'\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",0\r\n           ,''\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",0\r\n           ,''\r\n           " +
                            ",''\r\n           ,''\r\n           " +
                            ",''\r\n           " +
                            ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + "'\r\n           " +
                            ",'003'\r\n           " +
                            ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + "'\r\n           " +
                            ",'060'\r\n           " +
                            ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + "'\r\n           " +
                            ",'" + worksheet.Cells[row, 6].Value + "'\r\n           " +
                            ",null\r\n           " +
                            ",'N'\r\n           " +
                            ",null\r\n           " +
                            ",'1230'\r\n           " +
                            ",''\r\n           " +
                            ",1\r\n\t\t   )" +
                            "\r\nGO");
                        //}
                    }
                    System.IO.FileInfo file = new System.IO.FileInfo(directoryPath + "\\script.txt");
                    file.Directory.Create();
                    System.IO.File.WriteAllText(file.FullName, sb.ToString());
                }
                MessageBox.Show("Success! Pls open input folder and script.txt to view results");
            }
            catch(Exception ex) {
                MessageBox.Show(ex.ToString());
            }

        }
    }
}
