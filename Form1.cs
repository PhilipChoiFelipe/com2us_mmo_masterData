using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using MySqlConnector;
using System.Configuration;
using System.Linq;
using Dapper;


namespace com2us_mmo_masterData
{
    public partial class Form1 : Form
    {
        private OleDbConnection con = null;
        public Dictionary<String, DataTable> dictExcel = new Dictionary<string, DataTable>();
        private MySqlConnection _sqlConnection;
        public Form1()
        {
            InitializeComponent();
            this.Load += new EventHandler(Form1_Load);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string GameDBConnectionString = ConfigurationManager.AppSettings["DBConnection"];
                _sqlConnection = new MySqlConnection(GameDBConnectionString);
            }
            catch (Exception ex)
            {
                string ErrorMessage = ex.ToString();
                MessageBox.Show($"Failed to connect to database: {ErrorMessage}");
            }
        }

        public void getConnection(string fileName, string fileExt)
        {
            var conn = string.Empty;
            if (fileExt.CompareTo(".xls") == 0) 
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';";   
            else  
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;';";

            con = new OleDbConnection(conn);
        }
        
        public DataTable ReadSheet(String sheetName)
        {
            sheetName = TrimSheetName(sheetName);
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
            var idx = 0;
            foreach (DataRow row in infoTable.Rows)
            {
                var sheetName = row["TABLE_NAME"].ToString();
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

        String TrimSheetName(String sheetName)
        {
            if (sheetName.LastIndexOf('$') == sheetName.Length - 1)
            {
                sheetName = sheetName.Substring(0, sheetName.Length - 1);
            }

            return sheetName;

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;  
            var fileExt = string.Empty;  
            OpenFileDialog file = new OpenFileDialog(); 
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {  
                filePath = file.FileName;  
                fileExt = Path.GetExtension(filePath); 
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0) 
                {  
                    try
                    {
                        getConnection(filePath, fileExt);
                        
                        DataTable sheet = new DataTable();  
                        String[] sheetNameList=GetSheetName();
                        ReadData(sheetNameList);
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(sheetNameList);
                        
                        if (sheetNameList.Length >= 1)
                        {
                            sheet = ReadSheet(sheetNameList[0]);
                            comboBox1.SelectedIndex = 0;
                            dataGridView1.Visible = true;
                            dataGridView1.DataSource = sheet;
                        }
                 
                    } 
                    catch (Exception ex)
                    {  
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
            _sqlConnection.Open();
            foreach (KeyValuePair<string, DataTable> entry in dictExcel)
            {
                var sheetName = TrimSheetName(entry.Key);
                var paramList = "";
                var valueList = "";
                DynamicParameters parameter = new DynamicParameters();
                string[] columnNames = entry.Value.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                
                for (int i = 0; i < columnNames.Length; i++)
                {
                    paramList += $"{columnNames[i]},";
                    valueList += $"@{columnNames[i]},";
                    parameter.Add($"@{columnNames[i]}", entry.Value.Rows[0][i], direction: ParameterDirection.Input);
                }

                paramList = paramList.TrimEnd(',');
                valueList = valueList.TrimEnd(',');
                var sqlQuery = String.Format("INSERT {0} ({1}) VALUES({2}) ", sheetName, paramList, valueList);
                
                try
                {
                    var result = _sqlConnection.Execute(sqlQuery, parameter);
                    MessageBox.Show($"{sheetName} 마스터 데이터가 성공적으로 입력됨");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(String.Format("시트 에러 발생 {0}: {1}",sheetName, ex.ToString()));
                }
            }
        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable sheet = new DataTable();
            sheet = ReadSheet(comboBox1.Text.Trim());
            dataGridView1.DataSource=sheet;
        }
    }
}