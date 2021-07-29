using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;


namespace com2us_mmo_masterData
{
    public partial class Form1 : Form
    {
        public static DataTable dtExcel;
        public Form1()
        {
            InitializeComponent();
        }
        
        public DataTable ReadExcel(string fileName, string fileExt) {  
            string conn = string.Empty;  
            DataTable dtexcel = new DataTable();  
            if (fileExt.CompareTo(".xls") == 0) 
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";   
            else  
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; 
            using(OleDbConnection con = new OleDbConnection(conn)) {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [CharacterBaseData$]", con);
                    oleAdpt.Fill(dtexcel);
                }
                catch
                {
                    MessageBox.Show("CharacterBaseData Sheet Not Found.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }  
            }  
            return dtexcel;  
        }  


        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;  
            string fileExt = string.Empty;  
            OpenFileDialog file = new OpenFileDialog(); 
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {  
                filePath = file.FileName;  
                fileExt = Path.GetExtension(filePath); 
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0) {  
                    try {  
                        dtExcel = new DataTable();  
                        dtExcel = ReadExcel(filePath, fileExt); 
                        dataGridView1.Visible = true;  
                        dataGridView1.DataSource = dtExcel;  
                    } catch (Exception ex) {  
                        MessageBox.Show(ex.Message.ToString());  
                    }  
                } else {  
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }  
            }  
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            
        }
    }
}