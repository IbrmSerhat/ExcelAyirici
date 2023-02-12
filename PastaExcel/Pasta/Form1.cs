using System.Data;
using System.Data.OleDb;

namespace Pasta
{
    public partial class Form1 : Form
    {
        public List<DataTable> sehirler = new List<DataTable>();
        public DataTable tablo = new DataTable();
        public String connectionString = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void SehirleriAyristir(DataTable genelTablo, int sutun)
        {
            sehirler.Clear();
            comboBox2.Items.Clear();

            foreach (DataRow dr in genelTablo.Rows)
            {
                var filteredDataTables = sehirler.Where(dt => dt.TableName == dr[sutun].ToString().Trim());
                if (filteredDataTables.Any())
                {
                    filteredDataTables.First().ImportRow(dr);
                }
                else
                {
                    DataTable yeniTablo = genelTablo.Clone();
                    yeniTablo.TableName = dr[sutun].ToString();
                    yeniTablo.Rows.Clear();
                    yeniTablo.ImportRow(dr);
                    sehirler.Add(yeniTablo);
                    comboBox2.Items.Add(dr[sutun].ToString());
                }
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dosyalarý (*.xlsx)|*.xlsx";
            openFileDialog.Title = "Bir Excel dosyasý seçin";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    if (dt == null)
                    {
                        MessageBox.Show("Excel dosyasýnda hiçbir sayfa bulunamadý.");
                        return;
                    }

                    // Tüm Sheet Ýsimlerini Çekme
                    DataTable dtSayfaAdlari = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    // ComboBox'a Sheet Ýsimlerini Ekleme
                    foreach (DataRow row in dtSayfaAdlari.Rows)
                    {
                        string sheetName = row["TABLE_NAME"].ToString();
                        comboBox1.Items.Add(sheetName);
                    }

                    string sheetNamei = "11.02.2023$";
                    OleDbCommand command = new OleDbCommand("SELECT * FROM [" + comboBox1.Items[0] + "]" , connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);

                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);

                    foreach(DataColumn c in dataSet.Tables[0].Columns)
                    {
                        comboBox3.Items.Add(c.ColumnName);
                    }
                    tablo = dataSet.Tables[0];
                    dataGridView1.DataSource = dataSet.Tables[0];
                    SehirleriAyristir(dataSet.Tables[0], 0);
                    dataGridView2.DataSource = sehirler.First();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand commandd = new OleDbCommand("SELECT * FROM [" + comboBox1.Text + "]", connection);
                OleDbDataAdapter adapterr = new OleDbDataAdapter(commandd);

                DataSet dataSett = new DataSet();
                adapterr.Fill(dataSett);
                tablo = dataSett.Tables[0];
                dataGridView1.DataSource = dataSett.Tables[0];
                comboBox3.Items.Clear();
                foreach (DataColumn c in dataSett.Tables[0].Columns)
                {
                    comboBox3.Items.Add(c.ColumnName);
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var filteredDataTables = sehirler.Where(dt => dt.TableName == comboBox2.Text);
            if (filteredDataTables.Any())
            {
                dataGridView2.DataSource = filteredDataTables.First();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Text = comboBox3.SelectedIndex.ToString();
            SehirleriAyristir(tablo, comboBox3.SelectedIndex);
            if (sehirler.Any())
            {
                dataGridView2.DataSource = sehirler.First();
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            // Verileri kaydedeceðimiz Excel dosyasýnýn adý
            string fileName = "ExcelData.xlsx";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "|*.xlsx";
            saveFileDialog.Title = "dýþa aktar";
            saveFileDialog.FileName = "List";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog.FileName;


                // OLEDB baðlantýsý için connection string oluþturma
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'";

                // OLEDB connection nesnesi oluþturma
                OleDbConnection connection = new OleDbConnection(connectionString);

                DataTable dataTable = new DataTable();
                foreach (DataGridViewColumn column in dataGridView2.Columns)
                {
                    dataTable.Columns.Add(column.HeaderText);
                }

                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    DataRow dataRow = dataTable.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dataRow[cell.ColumnIndex] = cell.Value;
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // Excel dosyasýný oluþturma ve verileri ekleme
                try
                {
                    // OLEDB command nesnesi oluþturma
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;

                    // Baðlantý açma
                    connection.Open();

                    // Verileri kaydetmek için kullanýlacak sorgu
                    command.CommandText = "CREATE TABLE [DataTable] (";
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        command.CommandText += "[" + column.ColumnName + "] varchar(255),";
                    }
                    command.CommandText = command.CommandText.TrimEnd(',') + ")";

                    // Sorguyu çalýþtýrma
                    command.ExecuteNonQuery();

                    // Verileri kaydetmek için kullanýlacak sorgu
                    foreach (DataRow row in dataTable.Rows)
                    {
                        command.CommandText = "INSERT INTO [DataTable] VALUES (";
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            command.CommandText += "'" + row[column].ToString().Replace("'", "''") + "',";
                        }
                        command.CommandText = command.CommandText.TrimEnd(',') + ")";
                        command.ExecuteNonQuery();
                    }
                    MessageBox.Show("Veriler baþarýyla Excel dosyasýna aktarýldý.");
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Veriler aktarýlýrken hata oluþtu: " + ex.Message);
                }
                finally
                {
                    // Baðlantý kapatma
                    connection.Close();
                }
            }
            else
            {
                
            }
        }
    }
}

