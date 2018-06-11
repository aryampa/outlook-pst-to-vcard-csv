using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;
using System.Diagnostics;

using Microsoft.Office.Interop.Outlook;

using outlookObj = Microsoft.Office.Interop.Outlook;


namespace Outlook_TO_Vcf_CSV_EXCEL
{
    public partial class Form1 : Form
    {
        StringBuilder sBuilder;
        bool imageCreated = false;




        outlookObj.Application outlookApp;
        outlookObj.NameSpace ns;
        outlookObj.MAPIFolder contactsFolder;
        double contactsTotal, processed_Contacts, workPercentage;
        string out_path;

        string picturepath;
        //FileStream fs;
        //StreamWriter sWriter;

        string fullName;

        public Form1()
        {
            InitializeComponent();
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                fullName = "";
                out_path = "";
                sBuilder = new StringBuilder();
                outlookApp = new outlookObj.Application();
                ns = outlookApp.GetNamespace("MAPI");
                contactsFolder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                contactsTotal = contactsFolder.Items.Count;

                


            }

            catch (SystemException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void BtnLoad_Click(object sender, EventArgs e)
        {
            DialogResult result = folderDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                out_path = folderDialog.SelectedPath;
               // bgWorker.RunWorkerAsync(); // start background worker do_work routine



//#####################################################################################################################################

                try
                {
                    if (tbxFileName.Text == "")
                    {
                        throw new SystemException("File Name Not Provided!!!");
                    }

                    else
                    {

                        if (rbtnVcf.Checked) { ///// IF RBTN VCF


                           // fs = new FileStream(out_path + "\\"+tbxFileName .Text +".vcf", FileMode.Create, FileAccess.Write);
                            string fNam = "";
                            string MNan = "";
                            string LNam = "";

                           // sWriter = new StreamWriter(fs,Encoding.ASCII,100);


                            double total_contatcs = (double)contactsFolder.Items.Count;

                            double work_counter = 1;

                            sBuilder.Clear();

                            foreach (ContactItem contact_record in contactsFolder.Items)
                            {

                                lblProgressText.Text = work_counter.ToString() + " of " + contactsTotal.ToString();
                                lblProgressText.Update();

                                if (contact_record.FullName == null)
                                {
                                }

                                else
                                {

                                    fullName = contact_record.FullName.ToString();


                                    sBuilder.AppendLine("BEGIN:VCARD");
                                    sBuilder.AppendLine("VERSION:2.1");

                                    if (contact_record.FirstName != null) fNam = contact_record.FirstName.ToString();
                                    if (contact_record.MiddleName != null) MNan = contact_record.MiddleName.ToString();
                                    if (contact_record.LastName != null) LNam = contact_record.LastName.ToString();


                                    sBuilder.AppendLine("N:" + fNam + ";" + MNan + ";" + LNam);

                                    sBuilder.AppendLine("FN:" + fNam + " " + MNan + " " + LNam);



                                   

                                    if (contact_record.CompanyName != null) sBuilder.AppendLine("ORG:" + contact_record.CompanyName.ToString());
                                    if (contact_record.JobTitle != null) sBuilder.AppendLine("TITLE:" + contact_record.JobTitle.ToString());
                                    if (contact_record.MobileTelephoneNumber != null) sBuilder.AppendLine("TEL;CELL:" + contact_record.MobileTelephoneNumber.ToString());
                                    if (contact_record.BusinessTelephoneNumber != null) sBuilder.AppendLine("TEL;WORK:" + contact_record.BusinessTelephoneNumber.ToString());
                                    if (contact_record.Business2TelephoneNumber != null) sBuilder.AppendLine("TELL;WORK:" + contact_record.Business2TelephoneNumber.ToString());
                                    if (contact_record.OtherTelephoneNumber != null) sBuilder.AppendLine("TELL;WORK:" + contact_record.OtherTelephoneNumber.ToString());

                                    if (contact_record.HomeTelephoneNumber != null) sBuilder.AppendLine("TELL;HOME:" + contact_record.HomeTelephoneNumber.ToString());
                                    if (contact_record.Home2TelephoneNumber != null) sBuilder.AppendLine("TELL;HOME:" + contact_record.Home2TelephoneNumber.ToString());
                                    if (contact_record.Email1Address != null) sBuilder.AppendLine("EMAIL;WORK:" + contact_record.Email1Address.ToString());

                                    if (contact_record.Email2Address != null) sBuilder.AppendLine("EMAIL;HOME:" + contact_record.Email2Address.ToString());

                                    if (contact_record.HasPicture)
                                    {

                                        foreach (outlookObj.Attachment att in contact_record.Attachments)
                                        {

                                            if (att.DisplayName.ToString().Contains("ContactPicture"))
                                            {


                                                string picturepath = out_path + "\\" + contact_record.FullName.ToString() + "_" + contact_record.EntryID.ToString() + ".jpg";

                                                if (!File.Exists(picturepath))
                                                {
                                                    att.SaveAsFile(picturepath);

                                                    imageCreated = true;
                                                }

                                                if (imageCreated)
                                                {

                                                    string base64String_Image = "";

                                                    MemoryStream ms = new MemoryStream();

                                                    Image contactImg = Image.FromFile(picturepath);

                                                    contactImg.Save(ms, ImageFormat.Jpeg);

                                                    byte[] imageBytess = ms.ToArray();

                                                    base64String_Image = Convert.ToBase64String(imageBytess);

                                                    sBuilder.AppendLine("PHOTO;ENCODING=BASE64;JPEG:" + base64String_Image);

                                                    imageCreated = false;
                                                }

                                            }
                                        }

                                    }

                                    sBuilder.AppendLine("END:VCARD");



                                    processed_Contacts = work_counter;
                                    work_counter++;
                                }
                            }

                            
                           // sWriter.Write(sBuilder.ToString());

                          
                           //sWriter.Write("This is test test");
                           // fs.Flush();

                            TextWriter tw = File.CreateText(out_path + "\\" + tbxFileName.Text + ".vcf");
                            tw.Write(sBuilder.ToString());

                            tw.Close();

                            MessageBox.Show("Vcard Successfully Created!!!");

                            lblProgressText.Text = "";



                        }

                        else if (RbtnCsv.Checked) { /////// IF RBTN CSV


                        }
                    }
                }

                catch (SystemException ex)
                {
                    MessageBox.Show(ex.Message,"Error");
                    
                }


//#####################################################################################################################################

            }

            else MessageBox.Show("Output Directory Selection Canceled");







        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                Process[] running_Processes = Process.GetProcesses();
                foreach (Process active_process in running_Processes)
                {
                    if (active_process.ProcessName.ToUpper() == "OUTLOOK.EXE")
                    {
                        MessageBox.Show(active_process.ProcessName.ToUpper());
                        active_process.Close();
                    }
                }
            }

            catch (System.Exception ex) {
                MessageBox.Show("Error Closing OUTLOOK!! " + "ERROR: " + ex.Message.ToString());
                this.Close();

            }

        }
    }
}
