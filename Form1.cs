using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using MySqlConnector;
using System.Configuration;
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
        public static DataTable dtExcel;
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
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [NewCharacter$]", con);
                    oleAdpt.Fill(dtexcel);
                    Console.WriteLine(dtexcel.Columns);
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
            var filePath = string.Empty;  
            var fileExt = string.Empty;  
            OpenFileDialog file = new OpenFileDialog(); 
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {  
                filePath = file.FileName;  
                fileExt = Path.GetExtension(filePath); 
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0) 
                {  
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
           
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            _sqlConnection.Open();
            var SheetName = "NewCharacter";
            DynamicParameters parameter = new DynamicParameters();
            var sql = $"EXEC pInsert{SheetName}";
            for (int i = 0; i < dtExcel.Columns.Count; i++)
            {
                Console.WriteLine(dtExcel.Rows[0][i]);
                Console.WriteLine(dtExcel.Rows[1][i]);
                parameter.Add($"@{dtExcel.Rows[0][i]}", dtExcel.Rows[1][i], direction: ParameterDirection.Input);
            }
            parameter.Add("@InsertedId", dbType: DbType.String, direction: ParameterDirection.Output);
            
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
}