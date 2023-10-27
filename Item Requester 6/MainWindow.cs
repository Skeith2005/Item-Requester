using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.IO;
using System.Security.Authentication;
using Microsoft.Win32;
using LineItem;
using System.Diagnostics;
using System.Net;
namespace Item_Requester_6
{

    public partial class MainWindow : Form
    {
        BindingList<LineItem.LineItem> Items { get; set; }
        BindingList<String> Locations { get; set; }
        BindingList<String> MeasureUnits { get; set; }
        BindingList<LineItem.Names> NamesList { get; set; }
        public List<LineItem.UPC> UPCList { get; set; }
        public List<String> TactSKUList { get; set; }

        OpenFileDialog fileLoad;

        static SimpleAES enc = new SimpleAES();

        public string password = "Xenocide";
        public string localPath = "";
        public string localFile = "";
        public string savePath = "";
        public string reqDept = null;
        public string netUPCPath = @"\\10.107.54.188\Library\Item Requests\Program Files\UPC list.csv";
        public string currIndex;
        //public string localUPCPath = @"C:\Temp\Item Request Data\upc list.csv";
        public int nextTactSKU;
        public string pubFileName;
        public MailAddress ServiceRequest = new MailAddress("josh.shanabarger@pandphardware.com", "Josh Shanabarger");

        public bool unlistedName = false;
        public bool negGP = false;
        public bool checkUPC = false;
        public bool hasMatch = false;
        public bool firstChange = true;

        public MainWindow()
        {
            InitializeComponent();

            gdvItems.ReadOnly = false;
            gdvItems.AllowUserToDeleteRows = true;
            gdvItems.AutoGenerateColumns = false;
            gdvItems.AllowUserToResizeColumns = true;
            
            Items = new BindingList<LineItem.LineItem>();
            gdvItems.DataSource = Items;

            fileLoad = new OpenFileDialog();
            fileLoad.Filter = ".CSV|*.csv";

            NamesList = SetNameBox();
            cmbNames.DataSource = NamesList;
            cmbNames.DisplayMember = "FullName";

            Locations = new BindingList<String>();
            Locations.Add("Select Location");
            Locations.Add("Ace Hardware");
            Locations.Add("P & P Uniforms");

            cmbLocation.DataSource = Locations;
            cmbLocation.DisplayMember = "Name";

            MeasureUnits = new BindingList<String>();

            MeasureUnits.Add("EA - Each");
            MeasureUnits.Add("CS - Cases");
            MeasureUnits.Add("PK - Packs");
            MeasureUnits.Add("BX - Boxes");
            MeasureUnits.Add("RO - Rolls");
            MeasureUnits.Add("PT - Pints");
            MeasureUnits.Add("QT - Quarts");
            MeasureUnits.Add("GL - Gallons");
            
            cmbPUM.DataSource = MeasureUnits;
            //cmbPUM.DisplayMember = "Unit";

            cmbSUM.DataSource = MeasureUnits;
            //cmbSUM.DisplayMember = "Unit";

            DataColumnsSet(MeasureUnits);

        }
        
        #region Menu and Button Methods
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbNames.SelectedIndex == 0)
                {
                    MessageBox.Show("Name not selected. If your name is not listed, select \"Unlisted Name\" and add name in the Message field to be added to the list.", "Submission Error");
                    return;
                }

                if (cmbNames.SelectedIndex == 25 || cmbNames.Text == "Unlisted Name")
                {
                    unlistedName = true;
                }

                string[] nameSplit = cmbNames.Text.Split(' ');
                string name = nameSplit[0];


                string rootDir = @"\\10.107.54.188\Library\Item Requests\" + cmbLocation.Text;
                string saveName = String.Format("{0} - {1} - {2}.csv", cmbLocation.Text, name, DateTime.Now.ToString("yyyyMMdd"));

                if (Items.Count == 0)
                {
                    MessageBox.Show("Unable to submit because file is empty. Did you remember to click \"Add\" to add your item(s) to the list?", "Submission Error");
                    return;
                }
            }
            
            catch (Exception lenEx)
            {
                MessageBox.Show(lenEx.Message + lenEx.StackTrace);
                return;
            }

            FileInfo skuFile = new FileInfo(@"\\10.107.54.188\Library\Item Requests\Program Files\sku.txt");

            try
            {
                string file = WriteFile();

                if (file == null)
                    return;

                string tbxBody = string.Empty;
                string[] addArray = new string[6];
                int store;
                List<String> storeList = new List<String>();

                //shortMessage shortMsg = new shortMessage(/*unlistedName, cmbLocation.SelectedIndex*/);
                //shortMsg.ShowDialog();
                //tbxBody = shortMsg.bodyText;
                //addArray = shortMsg.storeAdd;

                string addConvert = ConvertStringArrayToString(addArray);
                string[] addWorker = addConvert.Split(',');
                
                for(int i = 0;i != addWorker.Length;i++)
                {
                    if(int.TryParse(addWorker[i], out store))
                    {
                        storeList.Add(store.ToString());
                    }
                }

                string result = String.Join(", ", storeList.ToArray());

                var client = new SmtpClient("smtp.office365.com", 587)
                {
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(@"josh.shanabarger@pandphardware.com", @"Xenocide123p@ss!"),
                    EnableSsl = true
                };



                MailMessage requesterMail = new MailMessage();
                requesterMail.IsBodyHtml = true;
               
                requesterMail.To.Add(ServiceRequest);
                requesterMail.From = new MailAddress(@"josh.shanabarger@pandphardware.com");
                requesterMail.Subject = file.Split('.')[0];

                requesterMail.Body = "An item request was submitted by " + cmbNames.Text + ". <br />" +
                                     "The following stores have been requested: " + result + "<br /><br />" +
                                     "Click <a href='mailto:" + GetReplyAddress(cmbNames.SelectedIndex) + "?subject=Item Request&cc=" + GetManagerAddress(cmbNames.SelectedIndex) + "'>here</a> to reply with SKUs.";

                if (cmbLocation.SelectedIndex != 2)
                {
                    requesterMail.Body += "<b><br />" + "Build department: " + reqDept + ".</b>";
                }

                if (tbxBody.Length != 0)
                    requesterMail.Body += "<b><i><br />" + "Attached Message: " + tbxBody + "</b></i>";

                if (cmbLocation.SelectedIndex != 2)
                {
                    foreach (var item in Items)
                    {
                        if (item.UPC.Length == 0)
                        {
                            requesterMail.Body += "<br />" + "This build contains items with no UPC. Confirm build with manager before finalizing.";
                            requesterMail.Priority = MailPriority.High;
                            break;
                        }
                    }
                }

                client.EnableSsl = true;

                try
                {
                    client.Send(requesterMail);
                }

                catch (SmtpException smtpEx)
                {
                    MessageBox.Show(smtpEx.Message + Environment.NewLine + smtpEx.StackTrace);
                    return;

                }
                 
                localPath = string.Empty;
                localFile = string.Empty;

                DialogResult exitResult = MessageBox.Show("Your submission was sent successfully! Would you like to exit?", "Hurrah!", MessageBoxButtons.YesNo);


                if (exitResult == DialogResult.Yes)
                {
                    Application.Restart();
                }

                else
                {
                    //System.Diagnostics.Process.Start(Application.ExecutablePath);
                    Application.Exit();
                }



            }
            catch (Exception mailEx)
            {
                MessageBox.Show(mailEx.Message + mailEx.StackTrace);
                return;
            }

        }
        
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbLocation.SelectedIndex != 2)
                {
                    foreach (var item in Items)
                    {
                        if (tbxDept.Text != item.Department)
                        {
                            MessageBox.Show("Only one department allowed per request. Send a request for each department needed. Thanks!");
                            return;
                        }
                    }
                }

                if (cmbLocation.Text == "Select Location")
                {
                    MessageBox.Show("Select location before adding items.");
                    return;
                }

                else
                {
                    reqDept = tbxDept.Text;
                }

                for (int i = 0; i < tbxMultiple.Text.Length; i++)
                {
                    if (!char.IsNumber(tbxMultiple.Text[i]))
                    {
                        MessageBox.Show("Your Multiple box input is incorrect. Make sure that it contains only numbers.", "Submissions Error");
                        return;
                    }
                }

                if (tbxDescription.Text.Contains(','))
                {
                    MessageBox.Show("Descriptions and MFG #s cannot contain commas. Please delete any commas.");
                    return;
                }

                if (tbxMFG.Text.Contains(','))
                {
                    MessageBox.Show("Descriptions and MFG #s cannot contain commas. Please delete any commas.");
                    return;
                }

                if (tbxDept.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a department for this item.", "Input Error");
                    return;
                }

                if (tbxDept.TextLength != 2)
                {
                    MessageBox.Show("The department field must contain two characters.", "Input Error");
                    return;
                }

                if (tbxClass.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a class for this item.", "Input Error");
                    return;
                }

                if (tbxClass.TextLength != 3)
                {
                    MessageBox.Show("The class field must contain three characters.", "Input Error");
                    return;
                }

                if (tbxFineline.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a fineline for this item.", "Input Error");
                    return;
                }

                if (tbxFineline.TextLength != 5)
                {
                    MessageBox.Show("The fineline field must contain five characters.", "Input Error");
                    return;
                }

                if (tbxRetail.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a retail cost for this item.", "Input Error");
                    return;
                }

                if (float.Parse(tbxGP.Text) < 0.00)
                {
                    MessageBox.Show("Your item has an Negative GP. Make sure your cost is BELOW your retail and re-add.");
                    return;
                }

                if (tbxMultiple.Text == string.Empty)
                {
                    MessageBox.Show("Please enter an order multiple for this item.", "Input Error");
                    return;
                }

                if (tbxDescription.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a Description for this item.", "Input Error");
                    return;
                }

                if (tbxMFG.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a Manufacturer Number for this item.", "Input Error");
                    return;
                }

                if (tbxDept.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a Department Code for this item.", "Input Error");
                    return;
                }

                if (tbxVendor.Text == string.Empty)
                {
                    DialogResult vendorCheck = MessageBox.Show("Vendor field empty. Is this intentional?", "Input Error", MessageBoxButtons.YesNo);
                    if (vendorCheck == DialogResult.Yes)
                    {
                        tbxVendor.Text = "VENDR";
                    }

                    else
                    {
                        return;
                    }
                }

                if (tbxRPL.Text == string.Empty)
                {
                    MessageBox.Show("Please enter a Replacement Cost for this item.", "Input Error");
                    return;
                }

                if (tbxUPC.Text == String.Empty && cmbLocation.Text != "P & P Uniforms")
                {
                    MessageBox.Show("UPC field is empty. If intentional, please enter \"NOUPC\" into the UPC field and add again.");
                    return;
                }

                if (tbxUPC.Text != String.Empty || tbxUPC.Text.Contains("UPC") || tbxUPC.Text.Contains("NO") && cmbLocation.Text != "P & P Uniforms")
                {
                    foreach (var item in Items)
                    {
                        if (tbxUPC.Text == item.UPC && tbxUPC.Text != "NOUPC")
                        {
                            MessageBox.Show("This UPC is already present in this item build. Check your UPCs and try again.");
                            return;
                        }
                    }
                }

                if (tbxDept.Text.Contains("T") && cmbLocation.Text != "P & P Uniforms")
                {
                    MessageBox.Show("All Tactical builds MUST be sent from the P&P Uniforms location. Please change your location and add again.", "Error");
                    return;
                }

                if (tbxUPC.Text.Length == 11 && tbxUPC.Text.Substring(0, 1) != "0")
                {
                    tbxUPC.Text = "0" + tbxUPC.Text;
                }

                currIndex = cmbLocation.SelectedIndex.ToString();
                firstChange = false;

                if (tbxUPC.Text.Contains("NO") || tbxUPC.Text.Contains("UPC"))
                {
                    tbxUPC.Text = String.Empty;
                }

                Items.Add(new LineItem.LineItem
                {
                    Description = tbxDescription.Text.ToUpper(),
                    MFG = tbxMFG.Text.ToUpper(),
                    Department = tbxDept.Text.ToUpper(),
                    Class = tbxClass.Text.ToUpper(),
                    Fineline = tbxFineline.Text.ToUpper(),
                    Vendor = tbxVendor.Text.ToUpper(),
                    RPLCost = tbxRPL.Text.ToUpper(),
                    Retail = tbxRetail.Text.ToUpper(),
                    Purch = cmbPUM.Text.Substring(0, 2).ToUpper(),
                    Stock = cmbSUM.Text.Substring(0, 2).ToUpper(),
                    Pack = tbxMultiple.Text.ToUpper(),
                    UPC = tbxUPC.Text.ToUpper(),
                    GrossProfit = Decimal.Parse(tbxGP.Text),
                });

                LineItem.LineItem test = new LineItem.LineItem();

                if ((bool)cbxCheck.Checked)
                {
                    tbxDescription.Text = string.Empty;
                    tbxMFG.Text = string.Empty;
                    tbxDept.Text = string.Empty;
                    tbxClass.Text = string.Empty;
                    tbxFineline.Text = string.Empty;
                    tbxVendor.Text = string.Empty;
                    tbxRPL.Text = string.Empty;
                    tbxRetail.Text = string.Empty;
                    tbxMultiple.Text = string.Empty;
                    tbxLocation.Text = string.Empty;
                }

                gdvItems.AutoResizeColumns();
            }

            catch (Exception addEx)
            {
                MessageBox.Show(addEx.Message + addEx.StackTrace);
                return;
            }
        }
        
        private void mnuClear(object sender, EventArgs e)
        {
            try
            {
                tbxDescription.Text = string.Empty;
                tbxMFG.Text = string.Empty;
                tbxDept.Text = string.Empty;
                tbxClass.Text = string.Empty;
                tbxVendor.Text = string.Empty;
                tbxRPL.Text = string.Empty;
                tbxRetail.Text = string.Empty;
                tbxMultiple.Text = string.Empty;
                tbxUPC.Text = string.Empty;
                cmbPUM.Text = "EA - Each";
                cmbSUM.Text = "EA - Each";
            }
            catch (Exception clearEx)
            {
                MessageBox.Show(clearEx.Message + clearEx.StackTrace);
                return;
            }
        }
        
        private void mnuOpen(object sender, EventArgs e)
        {
            string fileName = "";

            fileLoad.ShowDialog();

            fileName = fileLoad.FileName;

            if (fileName == "")
            {
                return;
            }

            using(StreamReader sr = new StreamReader(fileName))   
            {
                string line;
                Items.Clear();
                while ((line = sr.ReadLine()) != null)
                {
                    LineItem.LineItem itemLoad = new LineItem.LineItem();

                    try
                    {
                        string[] split = line.Split(',');
                        List<string> lstring = split.ToList();
                        itemLoad.Description = lstring[0];
                        itemLoad.MFG = lstring[1];
                        itemLoad.Department = lstring[2];
                        itemLoad.Class = lstring[3];
                        itemLoad.Fineline = lstring[4];
                        itemLoad.Vendor = lstring[5];
                        itemLoad.RPLCost = lstring[7];
                        itemLoad.Retail = lstring[8];
                        itemLoad.List = lstring[9];
                        itemLoad.DesiredGP = lstring[10];
                        itemLoad.Purch = lstring[11];
                        itemLoad.Stock = lstring[12];
                        itemLoad.Pack = lstring[13];
                        itemLoad.UPC = lstring[15];

                        Items.Add(itemLoad);
                        itemLoad = null;
                        lstring = null;
                    }

                    catch (ArgumentOutOfRangeException)
                    {
                        MessageBox.Show("Selected file is not formatted correctly. Please select a file that conforms to the format of this program.");
                        break;
                    }

                    finally
                    {
                        gdvItems.AutoResizeColumns();
                    }
                }
            }   
        }
                
        private void mnuNew(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Application.ExecutablePath);
            Application.Exit();
        }   
        
        private void mnuSave(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            string saveFilename = "";
            saveFile.DefaultExt = ".csv";
            saveFile.Filter = "(Comma-Delimited) .csv|*.csv";

            saveFile.ShowDialog();

            saveFilename = saveFile.FileName;

            using (StreamWriter sw = new StreamWriter(saveFilename))
            {

                foreach (var item in Items)
                {
                    sw.Write(String.Format("{0},", item.Description));
                    sw.Write(String.Format("{0},", item.MFG));
                    sw.Write(String.Format("{0},", item.Department));
                    sw.Write(String.Format("{0},", item.Class));
                    sw.Write(String.Format("{0},", item.Fineline));
                    sw.Write(String.Format("{0},", item.Vendor));
                    sw.Write(String.Format("{0},", item.Vendor));
                    sw.Write(String.Format("{0},", item.RPLCost));
                    sw.Write(String.Format("{0},", item.Retail));
                    sw.Write(String.Format("{0},", item.List));
                    sw.Write(String.Format("{0},", ""));
                    sw.Write(String.Format("{0},", item.Purch));
                    sw.Write(String.Format("{0},", item.Stock));
                    sw.Write(String.Format("{0},", item.Pack));
                    sw.Write(String.Format("{0},", ""));
                    sw.Write(String.Format("{0},", item.UPC));
                    sw.Write(String.Format("{0},", item.Pack));
                    sw.Write(String.Format("{0},", ""));
                    sw.Write(String.Format("{0},", "60901"));
                    sw.Write(String.Format("{0},", "42000000"));
                    sw.Write(String.Format("{0},", item.ProductCode));
                    if (cmbLocation.Text == "Inside Sales")
                    {
                        sw.Write(String.Format("{0},", item.MFG));
                    }

                    try
                    {
                        sw.WriteLine();
                    }
                    catch (Exception writeLineEx)
                    {
                        MessageBox.Show(writeLineEx.Message + writeLineEx.StackTrace);
                        break;
                    }

                }
            }

        }
        
        private void mnuExit(object sender, EventArgs e)
        {
            DialogResult exit = MessageBox.Show("Are you sure you want to exit?", "Please don't go...", MessageBoxButtons.YesNo);

            if (exit == System.Windows.Forms.DialogResult.Yes)
            {
                this.Close();
            }

        }

        private void mnuCache_Click(object sender, EventArgs e)
        {
            pwPrompt prompt = new pwPrompt();
            prompt.ShowDialog();

            string passedPassword = prompt.password;

            string wildDir = @"\\10.107.54.188\Library\Item Requests\Wildomar";
            string temeDir = @"\\10.107.54.188\Library\Item Requests\Temecula";
            string rvrsdDir = @"\\10.107.54.188\Library\Item Requests\Riverside";
            string PPUDir = @"\\10.107.54.188\Library\Item Requests\P & P Uniforms";
            string insideDir = @"\\10.107.54.188\Library\Item Requests\Inside Sales";

            if (passedPassword == password)
            {
                //string[] fileList = Directory.GetFiles(wildDir, "*.csv");

                List<String> fileList = Directory.GetFiles(wildDir, "*.csv").ToList();
                fileList.AddRange(Directory.GetFiles(temeDir, "*.csv").ToList());
                fileList.AddRange(Directory.GetFiles(rvrsdDir, "*.csv").ToList());
                fileList.AddRange(Directory.GetFiles(PPUDir, "*.csv").ToList());
                fileList.AddRange(Directory.GetFiles(insideDir, "*.csv").ToList());

                foreach (string csv in fileList)
                {
                    File.Delete(csv);
                }

                MessageBox.Show("Cache cleared.");


            }

            else if (passedPassword != password)
            {
                MessageBox.Show("Password incorrect.");
                return;
            }


        }
        #endregion

        #region Miscellaneous Methods

        private string WriteFile()
        {
            bool tactical = false;

            if (cmbNames.Text.Length == 0)
            {
                MessageBox.Show("Please enter your name.", "Submission Error");
                return null;
            }

            MessageBox.Show("Pick a location to store a local copy. If the program crashes or malfunctions during the submission, manually e-mail the local copy to Service Request.", "Please Read");

            SaveFileDialog swSaveDialog = new SaveFileDialog();
            swSaveDialog.DefaultExt = ".csv";
            swSaveDialog.Filter = "(Comma-Delimited) .csv|*.csv";
            
            swSaveDialog.ShowDialog();

            FileInfo localFileInfo = new FileInfo(swSaveDialog.FileName);

            localPath = localFileInfo.FullName;
            localFile = localFileInfo.Name;

            string fileName = String.Format("{0} - {1} - {2}.csv", cmbLocation.Text, cmbNames.Text.Substring(0, 1).ToUpper() + cmbNames.Text.Substring(1, cmbNames.Text.Length - 1).ToLower(), DateTime.Now.ToString("yyyyMMdd"));
            
            pubFileName = fileName;
            
            if (cmbLocation.SelectedIndex == 2)
            {
                tactical = true;
                nextTactSKU = GetNextTactSKU();
            }

            using (StreamWriter sw = new StreamWriter(localPath))
            {

                foreach (var item in Items)
                {
                    sw.Write(String.Format("{0},", item.Description));
                    sw.Write(String.Format("{0},", item.MFG));
                    sw.Write(String.Format("{0},", item.Department));
                    sw.Write(String.Format("{0},", item.Class));
                    sw.Write(String.Format("{0},", item.Fineline));
                    sw.Write(String.Format("{0},", item.Vendor));
                    sw.Write(String.Format("{0},", item.Vendor));
                    sw.Write(String.Format("{0},", item.RPLCost));
                    sw.Write(String.Format("{0},", item.Retail));
                    sw.Write(String.Format("{0},", item.List));
                    sw.Write(String.Format("{0},", ""));
                    sw.Write(String.Format("{0},", item.Purch));
                    sw.Write(String.Format("{0},", item.Stock));
                    sw.Write(String.Format("{0},", item.Pack));
                    sw.Write(String.Format("{0},", ""));
                    sw.Write(String.Format("{0},", item.UPC));
                    sw.Write(String.Format("{0},", item.Pack));
                    sw.Write(String.Format("{0},", ""));
                    sw.Write(String.Format("{0},", "60901"));
                    sw.Write(String.Format("{0},", "42000000"));
                    sw.Write(String.Format("{0},", item.ProductCode));
                    if (tactical)
                    {
                        sw.Write(String.Format("T{0},", nextTactSKU.ToString()));
                        nextTactSKU++;
                    }

                    try
                    {
                        sw.WriteLine();
                    }
                    catch (Exception writeLineEx)
                    {
                        MessageBox.Show(writeLineEx.Message + writeLineEx.StackTrace);
                        break;
                    }
                }
            }

            if (tactical)
            {
                using (Requester.Classes.UNCAccessWithCredentials unc = new Requester.Classes.UNCAccessWithCredentials())
                {
                    unc.NetUseWithCredentials(@"\\10.107.54.188\Library\Item Requests\Program Files\sku.txt", "library", "PPDOM", "Hardware123");

                    File.WriteAllText(@"\\10.107.54.188\Library\Item Requests\Program Files\sku.txt", nextTactSKU.ToString());

                    unc.Dispose();
                }
            }

            try
            {
                savePath = Path.Combine(@"\\10.107.54.188\Library\Item Requests\" + cmbLocation.Text, fileName);
                int count = 2;
                FileInfo saveInfo = new FileInfo(savePath);
                bool fileExist = File.Exists(savePath);

                while(fileExist)
                {
                    if (File.Exists(savePath))
                    {
                        //savePath = savePath.Replace(cmbLocation.Text + @"\", String.Format(@"{0}\({1}) ", cmbLocation.Text, count++));
                        if (count <= 2)
                        {
                            savePath = saveInfo.DirectoryName + "\\" + saveInfo.Name.Substring(0, saveInfo.Name.Length - 4) + " (" + count++ + ")" + saveInfo.Extension;
                        }

                        else if (count > 2)
                        {
                            savePath = saveInfo.DirectoryName + "\\" + saveInfo.Name.Substring(0, saveInfo.Name.Length - 8) + " (" + count++ + ")" + saveInfo.Extension;
                        }

                        saveInfo = new FileInfo(savePath);
                    }

                    if (!File.Exists(savePath))
                        fileExist = false;
                }

                try
                {
                    using (Requester.Classes.UNCAccessWithCredentials unc = new Requester.Classes.UNCAccessWithCredentials())
                    {
                        unc.NetUseWithCredentials(savePath, "library", "ppdom", "Hardware123");

                        localFileInfo.CopyTo(savePath, false);

                        unc.Dispose();
                    }
                }

                catch (Exception unc)
                {
                    MessageBox.Show(unc.Message + unc.StackTrace);
                }

            }

            catch (Exception fileEx)
            {
                MessageBox.Show("Crashing on Copy method: " + fileEx.Message + fileEx.StackTrace);
            }

            return fileName;
        }

        /// <summary>
        /// A method that sets the data columns for the DataGridView.
        /// </summary>
        ///
        public void DataColumnsSet(BindingList<String> MeasureUnits)
        {
            
            DataGridViewTextBoxColumn descCol = new DataGridViewTextBoxColumn();
            descCol.DataPropertyName = "Description";
            descCol.HeaderText = "Description";
            descCol.MaxInputLength = 32;
            descCol.Resizable = DataGridViewTriState.True;
            //descCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            gdvItems.Columns.Add(descCol);

            DataGridViewTextBoxColumn mfgCol = new DataGridViewTextBoxColumn();
            mfgCol.DataPropertyName = "MFG";
            mfgCol.HeaderText = "MFG #";
            mfgCol.MaxInputLength = 14;
            mfgCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(mfgCol);

            DataGridViewTextBoxColumn deptCol = new DataGridViewTextBoxColumn();
            deptCol.DataPropertyName = "Department";
            deptCol.HeaderText = "Department";
            deptCol.MaxInputLength = 2;
            deptCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(deptCol);

            DataGridViewTextBoxColumn classCol = new DataGridViewTextBoxColumn();
            classCol.DataPropertyName = "Class";
            classCol.HeaderText = "Class";
            classCol.MaxInputLength = 3;
            classCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(classCol);

            DataGridViewTextBoxColumn finelineCol = new DataGridViewTextBoxColumn();
            finelineCol.DataPropertyName = "Fineline";
            finelineCol.HeaderText = "Fineline";
            finelineCol.MaxInputLength = 5;
            finelineCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(finelineCol);

            DataGridViewTextBoxColumn vendorCol = new DataGridViewTextBoxColumn();
            vendorCol.DataPropertyName = "Vendor";
            vendorCol.HeaderText = "Vendor";
            vendorCol.MaxInputLength = 5;
            vendorCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(vendorCol);

            DataGridViewTextBoxColumn rplCol = new DataGridViewTextBoxColumn();
            rplCol.DataPropertyName = "RPLCost";
            rplCol.HeaderText = "Repl. Cost";
            rplCol.MaxInputLength = 8;
            rplCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(rplCol);

            DataGridViewTextBoxColumn retailCol = new DataGridViewTextBoxColumn();
            retailCol.DataPropertyName = "Retail";
            retailCol.HeaderText = "Retail";
            retailCol.MaxInputLength = 8;
            retailCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(retailCol);

            DataGridViewTextBoxColumn gpCol = new DataGridViewTextBoxColumn();
            gpCol.DataPropertyName = "DesiredGP";
            gpCol.HeaderText = "Desired GP";
            gpCol.MaxInputLength = 5;
            gpCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(gpCol);

            DataGridViewTextBoxColumn listCol = new DataGridViewTextBoxColumn();
            listCol.DataPropertyName = "List";
            listCol.HeaderText = "List";
            listCol.MaxInputLength = 8;
            listCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(listCol);

            DataGridViewTextBoxColumn purchUMCol = new DataGridViewTextBoxColumn();
            purchUMCol.DataPropertyName = "Purch";
            purchUMCol.HeaderText = @"Purch. U/M";
            purchUMCol.MaxInputLength = 2;
            purchUMCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(purchUMCol);

            DataGridViewTextBoxColumn stockUMCol = new DataGridViewTextBoxColumn();
            stockUMCol.DataPropertyName = "Stock";
            stockUMCol.HeaderText = @"Stock U/M";
            stockUMCol.MaxInputLength = 2;
            stockUMCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(stockUMCol);

            DataGridViewTextBoxColumn packCol = new DataGridViewTextBoxColumn();
            packCol.DataPropertyName = "Pack";
            packCol.HeaderText = "Multiple";
            packCol.MaxInputLength = 5;
            packCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(packCol);

            DataGridViewTextBoxColumn upcCol = new DataGridViewTextBoxColumn();
            upcCol.DataPropertyName = "UPC";
            upcCol.HeaderText = "UPC";
            upcCol.MaxInputLength = 14;
            upcCol.Resizable = DataGridViewTriState.True;
            gdvItems.Columns.Add(upcCol);


        }
        
        /// <summary>
        /// A method that passes the decrypted string of the Admin Username to server.
        /// </summary>
        /// 
        //private string "library"
        //{
        //    byte[] smtpUser = enc.Encrypt("LocalAdmin");
        //    string encName = enc.ByteArrToString(smtpUser);

        //    return enc.DecryptString(encName);
        //}

        ///// <summary>
        ///// A method that passes the decrypted string of the Admin Password to server.
        ///// </summary>
        ///// 
        //private string "Hardware123"
        //{
        //    byte[] smtpPW = enc.Encrypt("La123p@ss!");
        //    string encPW = enc.ByteArrToString(smtpPW);

        //    return enc.DecryptString(encPW);
        //}
       
        /// <summary>
        /// Generates reply address based on contents of Person text box.
        /// </summary>
        /// <returns></returns>
        private string GetReplyAddress(int nameIndex)
        {
            return NamesList.ElementAt(nameIndex).EmailAddress;
        }

        private string GetManagerAddress(int nameIndex)
        {
            if (NamesList.ElementAt(nameIndex).ManagerEmail == "None")
                return null;

            else
                return NamesList.ElementAt(nameIndex).ManagerEmail;

        }
        private BindingList<LineItem.Names> SetNameBox()
        {
            BindingList<LineItem.Names> list = new BindingList<LineItem.Names>();
            string users = @"\\10.107.54.188\Library\Item Requests\Program Files\users.csv";
            using (Requester.Classes.UNCAccessWithCredentials unc = new Requester.Classes.UNCAccessWithCredentials())
            {
                unc.NetUseWithCredentials(users, "library", "ppdom", "Hardware123");

                using (StreamReader sr = new StreamReader(users))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        LineItem.Names name = new LineItem.Names();

                        string[] split = line.Split(',');
                        List<string> lstring = split.ToList();
                        name.FullName = lstring[0];
                        name.EmailAddress = lstring[1];

                        list.Add(name);
                        name = null;
                        lstring = null;
                    }

                }

                return list;
            }
        }

        static string ConvertStringArrayToString(string[] array)
        {
            StringBuilder builder = new StringBuilder();
            foreach (string value in array)
            {
                builder.Append(value);
                builder.Append(", ");
            }
            return builder.ToString();
        }

        public float GetGrossProfit(float cost, float retail)
        {
            return ((retail - cost) / retail) * 100;
        }

        public int GetNextTactSKU()
        {
            using (Requester.Classes.UNCAccessWithCredentials enc = new Requester.Classes.UNCAccessWithCredentials())
            {
                enc.NetUseWithCredentials(@"\\10.107.54.188\Library\Item Requests\Program Files\sku.txt", "LocalAdmin", "ppdom", "Hardware123");

                using (StreamReader sr = File.OpenText(@"\\10.107.54.188\Library\Item Requests\Program Files\sku.txt"))
                {
                    return int.Parse(sr.ReadToEnd());
                }
            }
        }
        #endregion

        private void tbxRetail_Leave(object sender, EventArgs e)
        {
            if (!tbxRetail.Text.Contains('.'))
            {
                tbxRetail.Text += ".99";
            }

            if (tbxRPL.Text != string.Empty && tbxRetail.Text != string.Empty)
            {
                tbxGP.Text = GetGrossProfit(float.Parse(tbxRPL.Text), float.Parse(tbxRetail.Text)).ToString("n2");
            }
            
            else if (tbxRPL.Text != string.Empty && tbxRetail.Text == string.Empty)
            {
                float gpPrice = (-(100 * float.Parse(tbxRPL.Text) / (35 - 100)));
                //Easier To Read: (-(100 * COST) / GP - 100))

                if (!Math.Round(gpPrice, 2).ToString().Contains(".99"))
                {
                    gpPrice = float.Parse(gpPrice.ToString().Replace(gpPrice.ToString().Substring(gpPrice.ToString().Length - 3, 3), ".99"));
                }

                tbxRetail.Text = gpPrice.ToString();
            }
        }

        private void gdvItems_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (reqDept != "0" || reqDept != String.Empty && Items.Count == 0)
            {
                reqDept = String.Empty;
                firstChange = true;
            }
        }

        private void cmbLocation_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!firstChange)
            {
                MessageBox.Show("Can't change location after inputing items. Please close the requester and try again.");
                cmbLocation.SelectedIndex = int.Parse(currIndex);
                return;
            }
        }

        private void tbxRPL_Leave(object sender, EventArgs e)
        {
            if (!tbxRPL.Text.Contains('.'))
            {
                tbxRPL.Text += ".00";
            }
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }

            catch (IOException)
            {
                return true;
            }

            finally
            {
                if (stream != null)
                    stream.Close();
            }

            return false;
        }

        //private void CheckUPCData()
        //{
        //    if (File.Exists(localUPCPath))
        //    {
        //        DateTime lastLocalWrite = File.GetLastWriteTime(localUPCPath);
        //        DateTime lastServerWrite = File.GetLastWriteTime(netUPCPath);

        //        if (lastLocalWrite != lastServerWrite)
        //        {
        //            DialogResult dr = MessageBox.Show("Local UPC data is out of date. Update? (If updating is causing program issues, press \"No\")", "UPC Data Out of Date", MessageBoxButtons.YesNo);

        //            if (dr == System.Windows.Forms.DialogResult.No)
        //            {
        //                checkUPC = false;
        //            }

        //            if (dr == System.Windows.Forms.DialogResult.Yes)
        //            {
        //                File.Delete(localUPCPath);
        //                File.Copy(netUPCPath, localUPCPath);
        //                File.SetLastWriteTime(localUPCPath, DateTime.Now);
        //            }
        //        }
        //    }

        //    else if (!File.Exists(localUPCPath))
        //    {
        //        DialogResult dr = MessageBox.Show("No local UPC data exists. without this data, Item Requester can't utilize the UPC Check feature. Download data?", "UPC Data Not Found", MessageBoxButtons.YesNo);

        //        if (dr == System.Windows.Forms.DialogResult.No)
        //        {
        //            MessageBox.Show("UPC Check disabled.");
        //            checkUPC = false;
        //        }

        //        else
        //        {
        //            try
        //            {
        //                File.Copy(netUPCPath, localUPCPath, true);
        //                File.SetLastWriteTime(localUPCPath, DateTime.Now);
        //            }
        //            catch
        //            {
        //                MessageBox.Show("Data copy failed. Disabling UPC Check...", "Data Download Failed");
        //                checkUPC = false;
        //            }
        //        }
        //    }
        //}
    }
}