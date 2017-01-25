using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Kisters
{
    public partial class KistersUC : UserControl
    {
        bool boxWorldActivate = false;
        bool boxPartActivate = false;
        bool isKistersActive = false;
        string fileName;
        string nodesList = "";
        string standardNotesNode = "";
        string partNotesNode = "";
        string annotationNotesNode = "";
        string descriptionNode = "";
        List<string> nodes;
        List<string> units = new List<string>();
        string surface = "";
        string volume = "";
        string density = "";
        string dx = "";
        string dy = "";
        string dz = "";

        private string serverConn = "";
        SqlConnection connection;
        private string insertSpName;
        private string insertSpNotes;
        private string tableName = "";

        public void SetData(string serverConnF, string insertSpNameF, string insertSpNotesF, string tableNameF)
        {
            serverConn = serverConnF;
            insertSpName = insertSpNameF;
            insertSpNotes = insertSpNotesF;
            tableName = tableNameF;
        }
        public void SetFile(string fileNameExt)
        {
           
            string nodeName = "";
            string nodeType = "";            
            string imageString = "";
            string valuesString = "";

            //dispaly file            
            displayFile(fileNameExt);
            
            //calculate and get physical properties            
            for (var i = 0; i < nodes.Count; i++)
            {
                var result = getPhysicalproperties(nodes[i].ToString());
                nodeName = result.Item1;//name
                nodeType = result.Item2;//name
                surface = result.Item3;//surface
                volume = result.Item4;//volume
                density = result.Item5;//density

                if (checkBox1.Checked == true)
                {
                    dx = "1";
                    dy = "1";
                    dz = "1";
                }
                else
                {
                    //activate min box 
                    activateMinBox(nodes[i].ToString());

                    //get dx, dy, dz
                    var dResult = getDXDYDZOfNode();
                    dx = dResult.Item1;
                    dy = dResult.Item2;
                    dz = dResult.Item3;
                }
                //get picture
                imageString = getImage();

                //save all data in db
                valuesString = "'" + nodeName + "', ";
                valuesString = valuesString + "'" + surface + "', ";
                valuesString = valuesString + "'" + volume + "', ";
                valuesString = valuesString + "'" + density + "', ";

                valuesString = valuesString + "'" + dx + "', ";
                valuesString = valuesString + "'" + dy + "', ";
                valuesString = valuesString + "'" + dz + "', ";

                valuesString = valuesString + "'" + imageString + "'";

                insertSp(insertSpName, valuesString);
            }

            //how to get last ID in Model table???
            int modelID = getLastRecord(tableName);

            if (standardNotesNode != "")
            {
                getNote(modelID, standardNotesNode, "Standard note");
            }

            if (partNotesNode != "")
            {
                getNote(modelID, partNotesNode, "Part note");
            }

            if (annotationNotesNode != "")
            {
                getNote(modelID, annotationNotesNode, "Annotation note");
            }

            if (descriptionNode != "")
            {
                getNote(modelID, descriptionNode, "Material description");
            }
        }

        private void Start3DVS()
        {
            //Starts the ViewStation
            axK3DVSAX1.Start3DVS();
        }

        private void Terminate3DVS()
        {
            //Terminate the ViewStation
            axK3DVSAX1.Terminate();
        }

        private int getLastRecord(string tableName)
        {
            int lastID = 0;
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

                    cmd.CommandText = "SELECT TOP 1 ID FROM [dbo]." + tableName + " ORDER BY ID DESC";
                    //cmd.CommandText = "SELECT TOP 1 ID FROM [dbo].Model ORDER BY ID DESC";
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;

                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        lastID = Int32.Parse(reader[0].ToString());
                    }
                        
                    reader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

            return lastID;
        }

        private void getNote(int modelID, string noteNode, string noteType)
        {
            string sXmlCall = "<Call Method = 'SetSelectedNodes'>" + noteNode + "</Call>";
            string sXmlResponse = "";
            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //get all properties
            sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);

                foreach (XmlNode node in doc.SelectNodes("//Attribute"))
                {
                    string attrKey = node.Attributes["key"].Value;
                    string attrValue = node.Attributes["value"].Value.Substring(node.Attributes["value"].Value.IndexOf("|") + 1);
                    string valuesString = "";
                    valuesString = "'" + modelID + "','" + noteType + "','" + attrKey + "','" + attrValue + "'";
                    insertSp(insertSpNotes, valuesString);
                }
            }
            else
            {
            }

        }
        private void insertSp(string SpName, string valuesString)
        {

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(serverConn);
            using (connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();

                    SqlCommand cmd = new SqlCommand();
                    SqlDataReader reader;

                    cmd.CommandText = SpName;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@valuesString", valuesString);
                    cmd.Connection = connection;

                    reader = cmd.ExecuteReader();
                    reader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        public KistersUC()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Please, choose file");
                return;
            }

            //display file
            displayFile(textBox1.Text);

            //physical properties
            for (var i = 0; i < nodes.Count; i++)
            {
                var result = getPhysicalproperties(nodes[i].ToString());
            }

            ModifySelection("Deselect");
        }

        private void displayFile(string fileName)
        {
            nodes = new List<string>();

            //Always call RestoreView() before calls that will delete or create a view, e.g. "OpenFile", "New3DView", ...
            axK3DVSAX1.RestoreView();
                        
            //Create the call which should be executed            
            string sXmlCall = "<Call Method='OpenFile'><FileName>" + fileName + "</FileName></Call>";
            string sXmlResponse = "";
            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
            //Check if the call succeeded
            if (0 == iRet)
            {
                //Capture to display the view in the control
                axK3DVSAX1.CaptureView();
            }
            else
            {
                //an error has occurred
                //check sXmlResponse for more information
            }

            
            //===without this Kisters saves small pictures when run app without VS???
            var w = new Form() { Size = new Size(0, 0) };
            Task.Delay(TimeSpan.FromMilliseconds(100))
                .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
            MessageBox.Show(w, "", "");
            //==========================================

            //get structure
            sXmlCall = "<Call Method = 'GetStructure'/>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //get list of all nodes with Type = "Part"
            //<Node Id="1161" Name="HOUSING" Type="Part">
            //save all nodes in string
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(sXmlResponse);
            nodesList = "";
            int nodeTemp = 0;
            string partName = "";

            foreach (XmlNode node in doc.SelectNodes("//Node [@Type]"))
            {
                if (node.Attributes["Type"].Value == "Part")
                {
                    partName = node.Attributes["Name"].Value;
                    //we have to find node with "Ri_BrepModel" right after node with "Part" with the same name
                    //or with name = "MechanicalTool.2" ???

                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        if ((childNode.Attributes["Type"].Value == "Ri_BrepModel") && ((childNode.Attributes["Name"].Value == partName) || (childNode.Attributes["Name"].Value == "MechanicalTool.2")))
                        {
                            nodeTemp = Int32.Parse(childNode.Attributes["Id"].Value);
                            goto linkExit;
                        }

                        if (childNode.HasChildNodes)
                        {
                            foreach (XmlNode grandChildNode in childNode.ChildNodes)
                            {
                                if ((grandChildNode.Attributes["Type"].Value == "Ri_BrepModel") && ((grandChildNode.Attributes["Name"].Value == partName) || (grandChildNode.Attributes["Name"].Value == "MechanicalTool.2")))
                                {
                                    nodeTemp = Int32.Parse(grandChildNode.Attributes["Id"].Value);
                                    goto linkExit;
                                }
                            }
                        } 
                    }

                linkExit:                    
                    nodesList = nodesList + "<NodeId>" + nodeTemp.ToString() + "</NodeId>";
                    nodes.Add("<NodeId>" + nodeTemp.ToString() + "</NodeId>");
                }
            }

            //also get
            //< Node Id = "360" Name = "Standard Notes:" Type = "Ri_Set" />
            //< Node Id = "361" Name = "Part Notes:" Type = "Ri_Set" />
            //< Node Id = "362" Name = "Annotation Notes:" Type = "Ri_Set" />
            //< Node Id = "363" Name = "Material Description:" Type = "Ri_Set" />

            foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
            {
                if (node.Attributes["Name"].Value == "Standard Notes:")
                {
                    standardNotesNode = "<NodeId>" + Int32.Parse(node.Attributes["Id"].Value).ToString() + "</NodeId>";
                }

                if (node.Attributes["Name"].Value == "Part Notes:")
                {
                    partNotesNode = "<NodeId>" + Int32.Parse(node.Attributes["Id"].Value).ToString() + "</NodeId>";
                }

                if (node.Attributes["Name"].Value == "Annotation Notes:")
                {
                    annotationNotesNode = "<NodeId>" + Int32.Parse(node.Attributes["Id"].Value).ToString() + "</NodeId>";
                }

                if (node.Attributes["Name"].Value == "Material Description:")
                {
                    descriptionNode = "<NodeId>" + Int32.Parse(node.Attributes["Id"].Value).ToString() + "</NodeId>";
                }
            }
        }


        private Tuple<string, string, string, string, string> getPhysicalproperties(string currentNode)
        {
                string nodeName = "";
                string nodeType = "";

              string sXmlCall = "<Call Method = 'SetSelectedNodes'>" + currentNode + "</Call>";
              string sXmlResponse = "";
              int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

              //calculate physical properties            
              sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>ComputePhysicalProperties</SelectionModifier></Call>";
              sXmlResponse = "";
              iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

              //get all properties
              sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
              sXmlResponse = "";
              iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

              if (0 == iRet)
              {
                  XmlDocument doc = new XmlDocument();
                  doc.LoadXml(sXmlResponse);
                  // Console.WriteLine(sXmlResponse);

                  foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                  {
                      nodeName = node.Attributes["Name"].Value;
                  }

                  foreach (XmlNode node in doc.SelectNodes("//Node [@Type]"))
                  {
                      nodeType = node.Attributes["Type"].Value;
                  }

                  foreach (XmlNode node in doc.SelectNodes("//PhysicalProperties [@Surface]"))
                  {
                      label1.Text = "Surface = " + mmOrInch(node.Attributes["Surface"].Value, 2);
                      surface = node.Attributes["Surface"].Value;
                  }

                  foreach (XmlNode node in doc.SelectNodes("//PhysicalProperties [@Volume]"))
                  {
                      label2.Text = "Volume = " + mmOrInch(node.Attributes["Volume"].Value, 3);
                      volume = node.Attributes["Volume"].Value;
                  }

                  foreach (XmlNode node in doc.SelectNodes("//PhysicalProperties [@Density]"))
                  {
                      //label3.Text = "Density = " + node.Attributes["Density"].Value + " kg/dm\u00b3";
                      density = node.Attributes["Density"].Value;
                  }
              }
              else
              {
                  Console.WriteLine("GetNodeProperties error");
              }

            return Tuple.Create(nodeName, nodeType, surface, volume, density);
        }

        private string mmOrInch(string value, int pow)
        {
            string powString = "";

            if (pow == 1)
                powString = "";

            if (pow == 2)
                powString = "\u00b2";

            if (pow == 3)
                powString = "\u00b3";

            //return Math.Round(Convert.ToDouble(value.Replace(".", ",")), 2).ToString()   + " mm" + powString; //for RU
            return Math.Round(Convert.ToDouble(value), 2).ToString() + " mm" + powString; //for US
        }

        private void activateMinBox(string currentNode)
        {
            //first show all again
            String sXml = @"<Call Method='ShowAll'/>";
            ExecuteXml(sXml);

            //find and select node
            sXml = @"<Call Method = 'SetSelectedNodes'>" + currentNode + "</Call>";
            ExecuteXml(sXml);

            string sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>Isolate</SelectionModifier></Call>";
            string sXmlResponse = "";
            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //activate ComputeMinimalBoundingBox
            sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>ComputePartBoundingBox</SelectionModifier></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //deselect
            ModifySelection("Deselect");
        }

         
        private void button4_Click(object sender, EventArgs e)
        {
            //get image
            //Create the call which should be executed            
            string sXmlCall = "<Call Method='GetImage' Response='true'><ExportFormat2D>PNG</ExportFormat2D></Call>";
            //string sXmlResponse = "<Response Method = 'GetImage' Error = 'SUCCESS'><ExportFormat2D>PNG</ExportFormat2D><Image>Image in Base64</Image></Response>";
            string sXmlResponse = "";
            string imageString = "";
            //Image image;
            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
            //Check if the call succeeded
            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                XmlNode root = doc.FirstChild;

                //Display the contents of the child nodes.
                if (root.HasChildNodes)
                {
                    //last child contains Image
                    imageString = root.LastChild.InnerText;
                }

                byte[] bytes = Convert.FromBase64String(imageString);
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    pictureBox1.Image = Image.FromStream(ms);
                }
            }
            else
            {
                //an error has occurred
                //check sXmlResponse for more information
            }

            if (pictureBox1.Image == null)
                return;

            //save file
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|PNG Image|*.png";
            saveFileDialog1.Title = "Save an Image File";
            saveFileDialog1.FileName = fileName;
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.FileName != "")
            {
                System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
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

        private void button5_Click(object sender, EventArgs e)
        {
            if (boxWorldActivate == false)
            {
                //find and select node
                string sXmlCall = "<Call Method = 'SetSelectedNodes'>" + nodesList + "</Call>";
                string sXmlResponse = "";
                int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //clear selected node
                sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>Isolate</SelectionModifier></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //activate Measurement_BoundingBoxWorld
                sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>ComputeWorldBoundingBox</SelectionModifier></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                if (0 == iRet)
                {
                    boxWorldActivate = true;
                    boxPartActivate = false;
                }

                //deselect
                ModifySelection("Deselect");
            }
            else
            {
                //find and select node
                string sXmlCall = "<Call Method = 'SetSelectedNodes'>" + nodesList + "</Call>";
                string sXmlResponse = "";
                int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //deactivate Measurement_BoundingBoxWorld
                sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>Isolate</SelectionModifier></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
                if (0 == iRet)
                {
                    boxWorldActivate = false;
                    boxPartActivate = false;
                }

                //deselect
                ModifySelection("Deselect");
            }

            getDXDYDZ();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (boxPartActivate == false)
            {
                //find and select node
                String sXml = @"<Call Method = 'SetSelectedNodes'>" + nodesList + "</Call>";
                ExecuteXml(sXml);

                string sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>Isolate</SelectionModifier></Call>";
                string sXmlResponse = "";
                int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //activate ComputeMinimalBoundingBox
                sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>ComputeMinimalBoundingBox</SelectionModifier></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
                if (0 == iRet)
                {
                    boxWorldActivate = false;
                    boxPartActivate = true;
                }

                //deselect
                ModifySelection("Deselect");

            }
            else
            {
                //find and select node
                String sXml = @"<Call Method = 'SetSelectedNodes'>" + nodesList + "</Call>";
                ExecuteXml(sXml);

                string sXmlCall = "<Call Method = 'ModifySelection'><SelectionModifier>Isolate</SelectionModifier></Call>";
                string sXmlResponse = "";
                int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
                if (0 == iRet)
                {
                    boxWorldActivate = false;
                    boxPartActivate = false;
                }

                //deselect
                ModifySelection("Deselect");

            }
            getDXDYDZ();
        }

        private void getDXDYDZ()
        {
            //< Node Id = "405" Name = "DX=2065.6" Type = "Markup_Text"        
            string dxNode = "";
            string dyNode = "";
            string dzNode = "";

            ////==================DX
            string sXmlCall = "<Call Method = 'SearchNodes'><SearchString>Name = \"DX=*\"</SearchString></Call>";
            string sXmlResponse = "";
            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.Name == "NodeId")
                    {
                        dxNode = node.OuterXml;
                    }
                }
            }

            if (dxNode == "")
            {
                dx = "1";
                label4.Text = "DX = 1";
            }
            else
            {
                //select Node
                sXmlCall = "<Call Method = 'SetSelectedNodes'>" + dxNode + "</Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //get properties of selected Node
                sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
                if (0 == iRet)
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(sXmlResponse);

                    foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                    {
                        dx = node.Attributes["Name"].Value.Substring(3);
                        label4.Text = "DX = " + mmOrInch(node.Attributes["Name"].Value.Substring(3), 1);
                    }
                }
            }

            //==================DY
            sXmlCall = "<Call Method = 'SearchNodes'><SearchString>Name = \"DY=*\"</SearchString></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.Name == "NodeId")
                    {
                        dyNode = node.OuterXml;
                    }
                }
            }

            if (dyNode == "")
            {
                dy = "1";
                label5.Text = "DY = 1";
            }
            else
            {
                //select Node
                sXmlCall = "<Call Method = 'SetSelectedNodes'>" + dyNode + "</Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //get properties of selected Node
                sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
                if (0 == iRet)
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(sXmlResponse);

                    foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                    {
                        dy = node.Attributes["Name"].Value.Substring(3);
                        label5.Text = "DY = " + mmOrInch(node.Attributes["Name"].Value.Substring(3), 1);
                    }
                }
            }
            //==================DZ
            sXmlCall = "<Call Method = 'SearchNodes'><SearchString>Name = \"DZ=*\"</SearchString></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.Name == "NodeId")
                    {
                        dzNode = node.OuterXml;
                    }
                }
            }

            if (dzNode == "")
            {
                dz = "1";
                label6.Text = "DZ = 1";
            }
            else
            {
                //select Node
                sXmlCall = "<Call Method = 'SetSelectedNodes'>" + dzNode + "</Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

                //get properties of selected Node
                sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
                sXmlResponse = "";
                iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
                if (0 == iRet)
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(sXmlResponse);

                    foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                    {
                        dz = node.Attributes["Name"].Value.Substring(3);
                        label6.Text = "DZ = " + mmOrInch(node.Attributes["Name"].Value.Substring(3), 1);
                    }
                }
            }

            //deselect
            ModifySelection("Deselect");
        }

        private Tuple<string, string, string> getDXDYDZOfNode()
        {
            //< Node Id = "405" Name = "DX=2065.6" Type = "Markup_Text"        
            string dxNode = "";
            string dyNode = "";
            string dzNode = "";
            string dx = "";
            string dy = "";
            string dz = "";

            ////==================DX
            string sXmlCall = "<Call Method = 'SearchNodes'><SearchString>Name = \"DX=*\"</SearchString></Call>";
            string sXmlResponse = "";
            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.Name == "NodeId")
                    {
                        dxNode = node.OuterXml;
                    }
                }
            }

            //select Node
            sXmlCall = "<Call Method = 'SetSelectedNodes'>" + dxNode + "</Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //get properties of selected Node
            sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);

                foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                {
                    dx = node.Attributes["Name"].Value.Substring(3);
                    label4.Text = "DX = " + mmOrInch(node.Attributes["Name"].Value.Substring(3), 1);
                }
            }

            //==================DY
            sXmlCall = "<Call Method = 'SearchNodes'><SearchString>Name = \"DY=*\"</SearchString></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.Name == "NodeId")
                    {
                        dyNode = node.OuterXml;
                    }
                }
            }

            //select Node
            sXmlCall = "<Call Method = 'SetSelectedNodes'>" + dyNode + "</Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //get properties of selected Node
            sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);

                foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                {
                    dy = node.Attributes["Name"].Value.Substring(3);
                    label5.Text = "DY = " + mmOrInch(node.Attributes["Name"].Value.Substring(3), 1);
                }
            }
            //==================DZ
            sXmlCall = "<Call Method = 'SearchNodes'><SearchString>Name = \"DZ=*\"</SearchString></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.Name == "NodeId")
                    {
                        dzNode = node.OuterXml;
                    }
                }
            }

            //select Node
            sXmlCall = "<Call Method = 'SetSelectedNodes'>" + dzNode + "</Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);

            //get properties of selected Node
            sXmlCall = "<Call Method = 'GetNodeProperties'></Call>";
            sXmlResponse = "";
            iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);

                foreach (XmlNode node in doc.SelectNodes("//Node [@Name]"))
                {
                    dz = node.Attributes["Name"].Value.Substring(3);
                    label6.Text = "DZ = " + mmOrInch(node.Attributes["Name"].Value.Substring(3), 1);
                }
            }

            //deselect
            ModifySelection("Deselect");
            return Tuple.Create(dx, dy, dz);
        }

        private string getImage()
        {
            //get image
            //Create the call which should be executed            
            string sXmlCall = "<Call Method='GetImage' Response='true'><ExportFormat2D>PNG</ExportFormat2D></Call>";
            //string sXmlResponse = "<Response Method = 'GetImage' Error = 'SUCCESS'><ExportFormat2D>PNG</ExportFormat2D><Image>Image in Base64</Image></Response>";
            string sXmlResponse = "";
            string imageString = "";

            int iRet = axK3DVSAX1.ExecuteApiCall(sXmlCall, ref sXmlResponse);
            //Check if the call succeeded
            if (0 == iRet)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(sXmlResponse);
                XmlNode root = doc.FirstChild;

                //Display the contents of the child nodes.
                if (root.HasChildNodes)
                {
                    //last child contains Image
                    imageString = root.LastChild.InnerText;
                }
            }

            return imageString;
        }

        //from Kisters
        bool ExecuteXml(String sXml, String response = null)
        {
            String sXmlResponse = "";

            //Execute the xml call
            int iError = axK3DVSAX1.ExecuteApiCall(sXml, ref sXmlResponse);

            if (0 != sXmlResponse.Length)
            {
                if (null != response)
                {
                    response = sXmlResponse;
                }
            }

            return 0 == iError;
        }

        private void ModifySelection(String sModifier)
        {
            String sXml = @"<Call Method='ModifySelection'><SelectionModifier>" + sModifier + @"</SelectionModifier></Call>";
            ExecuteXml(sXml);
        }
            
        private void KistersUC_Load(object sender, EventArgs e)
        {
            Start3DVS();
        }
    }
}
