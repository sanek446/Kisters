using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Kisters
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //run Kisters3D
                kistersUC1.Start3DVS();
                
                foreach (String file in openFileDialog1.FileNames)
                {
                    Console.WriteLine(DateTime.Now);
                    Console.WriteLine(file);
                    kistersUC1.SetFile(file);
                    Console.WriteLine(DateTime.Now);
                }

                //kill Kisters3D
                kistersUC1.Terminate3DVS();
                MessageBox.Show("Done");                        
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string connectionString = "Server = (LocalDB)\\MSSQLLocalDB; AttachDbFilename=|DataDirectory|ModelsDB.mdf;Integrated Security = true";
            string insertSP = "insertSP";
            string insertSPNotes = "insertSPNotes";
            string tableName = "Model";

            kistersUC1.SetData(connectionString, insertSP, insertSPNotes, tableName);
        }


        private void button3_Click(object sender, EventArgs e)
        {
            selectFromBD();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "")
            {
                MessageBox.Show("Fill ID");
                return;
            }
            string serverConn = "Server = (LocalDB)\\MSSQLLocalDB; AttachDbFilename=|DataDirectory|ModelsDB.mdf;Integrated Security = true";
            SqlConnection connection;

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(serverConn);
            using (connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();

                    SqlCommand cmd = new SqlCommand();
                    SqlDataReader reader;

                    cmd.CommandText = "SELECT Image FROM [dbo].Model where ID = " + textBox1.Text;
                    cmd.CommandType = CommandType.Text;                    
                    cmd.Connection = connection;

                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        /*   Byte[] data = new Byte[0];
                           data = (Byte[])(reader[0]);
                           MemoryStream mem = new MemoryStream(data);
                           pictureBox1.Image = Image.FromStream(mem);*/

                        byte[] bytes = Convert.FromBase64String(reader[0].ToString());
                        using (MemoryStream ms = new MemoryStream(bytes))
                        {
                            pictureBox1.Image = Image.FromStream(ms);
                        }

                        //save file
                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                        saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|PNG Image|*.png";
                        saveFileDialog1.Title = "Save an Image File";
                        //saveFileDialog1.FileName = fileName;
                        saveFileDialog1.ShowDialog();

                        // If the file name is not an empty string open it for saving.
                        if (saveFileDialog1.FileName != "")
                        {
                            FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                            switch (saveFileDialog1.FilterIndex)
                            {
                                case 1:
                                    this.pictureBox1.Image.Save(fs,
                                       System.Drawing.Imaging.ImageFormat.Jpeg);
                                    break;
                                case 2:
                                    this.pictureBox1.Image.Save(fs,
                                       System.Drawing.Imaging.ImageFormat.Bmp);
                                    break;
                                case 3:
                                    this.pictureBox1.Image.Save(fs,
                                       System.Drawing.Imaging.ImageFormat.Png);
                                    break;
                            }

                            fs.Close();
                        }
                    }                 

                    reader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        
        private void button5_Click(object sender, EventArgs e)
        {
            string serverConn = "Server = (LocalDB)\\MSSQLLocalDB; AttachDbFilename=|DataDirectory|ModelsDB.mdf;Integrated Security = true";
            SqlConnection connection;

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(serverConn);
            using (connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();

                    SqlCommand cmd = new SqlCommand();
                    SqlDataReader reader;

                    cmd.CommandText = "DELETE FROM [dbo].Model";
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;

                    reader = cmd.ExecuteReader();
                    reader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

            builder = new SqlConnectionStringBuilder(serverConn);
            using (connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();

                    SqlCommand cmd = new SqlCommand();
                    SqlDataReader reader;

                    cmd.CommandText = "DELETE FROM [dbo].Notes";
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;

                    reader = cmd.ExecuteReader();
                    reader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

            selectFromBD();
        }


        private void selectFromBD()
        {
            string connectionString = "Server = (LocalDB)\\MSSQLLocalDB; AttachDbFilename=|DataDirectory|ModelsDB.mdf;Integrated Security = true";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlDataAdapter ad = new SqlDataAdapter("SELECT ID, Name, Surface AS 'Surface [mm\u00b2]', Volume AS 'Volume [mm\u00b3]', dx AS 'DX [mm]', dy AS 'DY [mm]', dz AS 'DZ [mm]' FROM [dbo].Model", conn))
                {
                    DataTable dt = new DataTable();
                    ad.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
            }

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlDataAdapter ad = new SqlDataAdapter("SELECT ID, ModelID, TypeOfNote, AttributeCode, AttributeText FROM [dbo].Notes", conn))
                {
                    DataTable dt = new DataTable();
                    ad.Fill(dt);
                    dataGridView2.DataSource = dt;
                }
            }
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string connectionString = "Server = (LocalDB)\\MSSQLLocalDB; AttachDbFilename=|DataDirectory|ModelsDB.mdf;Integrated Security = true";

            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string selectedID = row.Cells[0].Value.ToString();

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    using (SqlDataAdapter ad = new SqlDataAdapter("SELECT ID, ModelID, TypeOfNote, AttributeCode, AttributeText FROM [dbo].Notes WHERE ModelID = '" + selectedID + "'", conn))
                    {
                        DataTable dt = new DataTable();
                        ad.Fill(dt);
                        dataGridView2.DataSource = dt;
                    }
                }

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "export models.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {                
                ToCsV(dataGridView1, sfd.FileName);
            }

            sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "export notes.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView2, sfd.FileName);
            }
        }

        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }
    }
}
