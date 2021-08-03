using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.Logging;


namespace com2us_mmo_masterData
{
    public partial class Form1 : Form
    {
        private OleDbConnection con = null;
        public Dictionary<String, DataTable> dictExcel = new Dictionary<string, DataTable>();
        public Form1()
        {
            InitializeComponent();
        }

        public void getConnection(string fileName, string fileExt)
        {
            string conn = string.Empty;
            if (fileExt.CompareTo(".xls") == 0) 
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";   
            else  
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';";

            con = new OleDbConnection(conn);
        }


        
        public DataTable ReadSheet(String sheetName) 
        {
            if (sheetName.LastIndexOf('$') == sheetName.Length - 1)
                sheetName = sheetName.Substring(0, sheetName.Length - 1);
            con.Open();
            DataTable dtexcel = new DataTable();
    
            OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + sheetName + "$]", con);
            oleAdpt.Fill(dtexcel);
            con.Close();
            return dtexcel;
        }  
        
        public String[] GetSheetName()
        {
            DataTable infoTable = null;
            
            con.Open();
            infoTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            String[] sheet = new String[infoTable.Rows.Count];
            int idx = 0;
            foreach (DataRow row in infoTable.Rows)
            {
                String sheetName = row["TABLE_NAME"].ToString();
                sheet[idx++]=sheetName;
            }
            con.Close();
            
            return sheet;
        }

        void ReadData(String[] sheetNameList)
        {
            for (int i = 0; i < sheetNameList.Length; i++)
            {
                String k = sheetNameList[i];
                DataTable v = ReadSheet(k);
                dictExcel.Add(k,v);
            }
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
                    try
                    {
                        getConnection(filePath, fileExt);
                        
                        dtExcel = new DataTable();  
                        String[] sheetNameList=GetSheetName();
                        ReadData(sheetNameList);
                        comboBox1.Items.AddRange(sheetNameList);
                        
                        if (sheetNameList.Length >= 1)
                        {
                            dtExcel = ReadSheet(sheetNameList[0]);
                            comboBox1.SelectedIndex = 0;
                            dataGridView1.Visible = true;
                            dataGridView1.DataSource = dtExcel;
                        }
                 
                    } 
                    catch (Exception ex) {  
                        MessageBox.Show(ex.Message.ToString());  
                    }  
                }
                else 
                {  
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable sheet = new DataTable();
            sheet = ReadSheet(comboBox1.Text.Trim());
            dataGridView1.DataSource=sheet;
        }
    }
}