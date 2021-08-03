using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.Logging;
using MySqlConnector;
using System.Configuration;
using System.Linq;
using Dapper;
using System.Runtime.InteropServices;
//PlantUML
//TODO: 
/*
 * Datatable 배열로 만들기
 * 데이터 오류 생기면 오류 정보 return
 * api server에서 마스터 데이터 불러오기
 * game server에서 db manager로 가져오기 
 */
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
            AllocConsole();
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
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        public void getConnection(string fileName, string fileExt)
        {
            string conn = string.Empty;
            if (fileExt.CompareTo(".xls") == 0) 
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;TypeGuessRows=0;';";   
            else  
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;';";

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
            var filePath = string.Empty;  
            var fileExt = string.Empty;  
            OpenFileDialog file = new OpenFileDialog(); 
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {  
                filePath = file.FileName;  
                fileExt = Path.GetExtension(filePath); 
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0) {  
                    try
                    {
                        getConnection(filePath, fileExt);
                        
                        DataTable sheet = new DataTable();  
                        String[] sheetNameList=GetSheetName();
                        ReadData(sheetNameList);
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
                DynamicParameters parameter = new DynamicParameters();
                var sheetName = entry.Key.Substring(0, entry.Key.Length - 1);
                var sql = $"pInsert{sheetName}";
                Console.WriteLine(sql);
                string[] columnNames = entry.Value.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                for (int i = 0; i < columnNames.Length; i++)
                {
                    Console.WriteLine($"in{columnNames[i]}");
                    Console.WriteLine(entry.Value.Rows[0][i]);
                    Console.WriteLine(entry.Value.Rows[0][i].GetType());

                    parameter.Add($"@in{columnNames[i]}", entry.Value.Rows[0][i], direction: ParameterDirection.Input);
                }
                parameter.Add("@InsertedId", dbType: DbType.Int32, direction: ParameterDirection.Output);
            
                try
                {
                    _sqlConnection.Execute(sql, parameter, commandType: CommandType.StoredProcedure);
                    var insertedId = parameter.Get<int>("@InsertedId");
                    Console.WriteLine($"Inserted ID: {insertedId}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
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