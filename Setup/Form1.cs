using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using File = System.IO.File;
using System.Security.AccessControl;

namespace Setup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection SourceConnection = new SqlConnection("Data Source =" + Environment.MachineName + @"\SQLEXPRESS; Initial Catalog = master; Persist Security Info = True; User ID =sa ; Password = P@$$w0rd;");

        DriveInfo[] Suruculer = DriveInfo.GetDrives();
        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (DriveInfo d in Suruculer)
            {
                if (d.IsReady == true)
                {
                    comboBox1.Items.Add(d.Name);
                }
            }
            textBox2.Text = Environment.MachineName + " / " + Environment.UserName;
            comboBox1.SelectedIndex = 0;
            textBox1.Text = comboBox1.Text + "MyApp";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DosyaYoluTanımlama();
        }

      
        #region İnteraktive Metodlar
        public void ExtractFile(string sourceArchive, string destination)
        {
            string zPath = "7za.exe";
            try
            {
                ProcessStartInfo pro = new ProcessStartInfo();
                pro.WindowStyle = ProcessWindowStyle.Hidden;
                pro.FileName = zPath;
                pro.Verb = "runas";
                pro.Arguments = string.Format("x \"{0}\" -y -o\"{1}\"", sourceArchive, destination);
                Process x = Process.Start(pro);
                x.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                listBox1.Items.Add(ex);
                MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }
        }
        public static void AddFileSecurity(string fileName, string account, FileSystemRights rights, AccessControlType controlType)
        {
            FileSecurity fSecurity = File.GetAccessControl(fileName);
            fSecurity.AddAccessRule(new FileSystemAccessRule(account, rights, controlType));
            File.SetAccessControl(fileName, fSecurity);
        }
        public void database_attach(string db, string mdf, string ldf)
        {
            try
            {
                string sorgu = "use master " +
                "create database [" + db + "] on " +
                "(filename=N'" + mdf + "')," +
                "(filename=N'" + ldf + "')" +
                " for attach";
                if (SourceConnection.State == System.Data.ConnectionState.Closed)
                {
                    SourceConnection.Open();
                }
                SqlCommand attach = new SqlCommand(sorgu, SourceConnection);
                attach.ExecuteScalar();
                SourceConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                listBox1.Items.Add(ex);
                MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }
        }
        private void CreateShortcut()
        {
            try
            {
                listBox1.Items.Add("Masaüstü Kısayolu Oluşturuluyor....");
                if (Is64Bit())
                {
                    //64Bit windows için
                    object shDesktop = (object)"Desktop";
                    WshShell shell = new WshShell();
                    string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\BilsetNET.lnk";
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
                    shortcut.Description = "BilsetNET Kısayolu";
                    shortcut.TargetPath = textBox1.Text + @"\BilsetEntegre\BilsetEntegre.exe";
                    shortcut.Save();
                }
                else
                {
                    //32 Bit windows için
                    object shDesktop = (object)"Desktop";
                    WshShellClass shell = new WshShellClass();
                    string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\BilsetNET.lnk";
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
                    shortcut.Description = "BilsetNET Kısayolu";
                    shortcut.TargetPath = textBox1.Text + @"\BilsetEntegre\BilsetEntegre.exe";
                    shortcut.Save();
                }
                listBox1.Items.Add("Masaüstü Kısayolu Oluşturuldu.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                listBox1.Items.Add(ex);
                MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }

        }
        void DosyaYoluTanımlama()
        {
            try
            {
                string DosyaYolu = textBox1.Text;
                StringBuilder SeciliDrive = new StringBuilder(DosyaYolu);
                for (int i = 0; i < 2; i++)
                {
                    SeciliDrive[i] = Convert.ToChar(comboBox1.Text.ToCharArray(i, 1).GetValue(0).ToString());
                }
                DosyaYolu = SeciliDrive.ToString();
                textBox1.Text = DosyaYolu;
            }
            catch
            {

            }
        }

        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWow64Process([In] IntPtr hProcess, [Out] out bool lpSystemInfo);

        public bool Is64Bit()
        {
            bool retVal;
            IsWow64Process(Process.GetCurrentProcess().Handle, out retVal);

            return retVal;
        }
        #endregion

        #region Özet Metodlar
        void FireWallnTCPPort()
        {
            Process RunSetSQL = new Process();
            ProcessStartInfo SetSQLnInfo = new ProcessStartInfo();
            SetSQLnInfo.FileName = Application.StartupPath + @"\SetSQLTCPPort\SetSQLTCPPort.exe";
            SetSQLnInfo.Verb = "runas";
            RunSetSQL.StartInfo = SetSQLnInfo;
            RunSetSQL.Start();
            RunSetSQL.WaitForExit();
        }
        void SQLSETUP()
        {
            if (checkBox1.Checked == false)
            {
                progressBar1.Value += 80;
                return;
            }
            string DosyaYolu64 = Application.StartupPath + @"\SQL2014-x64\Setup.exe";
            string DosyaYolu32 = Application.StartupPath + @"\SQL2014-x86\Setup.exe";
            Process SQLEXTRACT = new Process();
            ProcessStartInfo ExtractionInfo = new ProcessStartInfo();
            Process SQLSetup = new Process();
            ProcessStartInfo SetupInfo = new ProcessStartInfo();
            if (Is64Bit())
            {
                if (!File.Exists(DosyaYolu64))
                {
                    try
                    {
                        ExtractionInfo.FileName = Application.StartupPath + @"\SQL2014-x64.exe";
                        ExtractionInfo.Arguments = ExtractionInfo.FileName + @" /u /x:" + Application.StartupPath + @"\SQL2014-x64";
                        ExtractionInfo.Verb = "runas";
                        SQLEXTRACT.StartInfo = ExtractionInfo;
                        SQLEXTRACT.Start();
                        progressBar1.Value += 5;
                        listBox1.Items.Add("SQL Server 2014(64-Bit) Setup dosyaları '" + Application.StartupPath + @"\SQL2014-x64' konumuna çıkarılıyor...");
                        SQLEXTRACT.WaitForExit();
                        progressBar1.Value += 5;
                        listBox1.Items.Add("SQL Server 2014(64-Bit) Setup dosyaları çıkarıldı.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        listBox1.Items.Add(ex);
                        MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                        return;
                    }

                }
                else
                {
                    progressBar1.Value += 5;
                    progressBar1.Value += 5;
                    listBox1.Items.Add("SQL Server 2014(64-Bit) Dosyaları zaten çıkartılmış, SQL kurulumuna geçiliyor.");
                }
                try
                {

                    SetupInfo.FileName = Application.StartupPath + @"\SQL2014-x64\Setup.exe";
                    SetupInfo.Arguments = SetupInfo.FileName + @" /ConfigurationFile=" + Application.StartupPath + @"\ConfigFile-x64.ini /Medialayout=Advanced /SAPWD=""P@$$w0rd""";
                    SetupInfo.Verb = "runas";
                    SQLSetup.StartInfo = SetupInfo;
                    SQLSetup.Start();
                    progressBar1.Value += 35;
                    listBox1.Items.Add("SQL Server 2014(64-Bit) yükleniyor... Bu işlem biraz zaman alabilir, lütfen bekleyiniz....");
                    SQLSetup.WaitForExit();
                    progressBar1.Value += 35;
                    listBox1.Items.Add("SQL Server 2014(64-Bit) yüklendi.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    listBox1.Items.Add(ex);
                    MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }
            }
            else
            {
                if (!File.Exists(DosyaYolu32))
                {
                    try
                    {
                        ExtractionInfo.FileName = Application.StartupPath + @"\SQL2014-x86.exe";
                        ExtractionInfo.Arguments = ExtractionInfo.FileName + @" /u /x:" + Application.StartupPath + @"\SQL2014-x86";
                        ExtractionInfo.Verb = "runas";
                        SQLEXTRACT.StartInfo = ExtractionInfo;
                        SQLEXTRACT.Start();
                        progressBar1.Value += 5;
                        listBox1.Items.Add("SQL Server 2014(32-Bit) Setup dosyaları '" + Application.StartupPath + @"\SQL2014-x86' konumuna çıkarılıyor...");
                        SQLEXTRACT.WaitForExit();
                        progressBar1.Value += 5;
                        listBox1.Items.Add("SQL Server 2014(32-Bit) Setup dosyaları çıkarıldı.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        listBox1.Items.Add(ex);
                        MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                        return;
                    }

                }
                else
                {
                    progressBar1.Value += 5;
                    listBox1.Items.Add("SQL Server 2014(32-Bit) Dosyaları zaten çıkartılmış, direkt setupa geçiliyor.");
                }
                try
                {
                    SetupInfo.FileName = Application.StartupPath + @"\SQL2014-x86\Setup.exe";
                    SetupInfo.Arguments = SetupInfo.FileName + @" /ConfigurationFile=" + Application.StartupPath + @"\ConfigFile-x86.ini /Medialayout=Advanced /SAPWD=""P@$$w0rd""";
                    SetupInfo.Verb = "runas";
                    SQLSetup.StartInfo = SetupInfo;
                    SQLSetup.Start();
                    progressBar1.Value += 35;
                    listBox1.Items.Add("SQL Server 2014(32-Bit) yükleniyor... Bu işlem biraz zaman alabilir, lütfen bekleyiniz....");
                    SQLSetup.WaitForExit();
                    progressBar1.Value += 35;
                    listBox1.Items.Add("SQL Server 2014(32-Bit) yüklendi.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    listBox1.Items.Add(ex);
                    MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }
            }
        }
        void AppKurulumu()
        {
            progressBar1.Value += 10;
            listBox1.Items.Add("App yükleniyor... Lütfen bekleyiniz....");
            if (!File.Exists(textBox1.Text + @"BilsetEntegre\BilsetEntegre.exe"))
            {
                ExtractFile(Application.StartupPath + @"\BilsetNET.7z", textBox1.Text);
                progressBar1.Value += 10;
                listBox1.Items.Add("App yüklendi.");
            }
            else
            {
                if (MessageBox.Show("App zaten yüklü güncellemek ister misiniz ?", "App Güncellemesi", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    listBox1.Items.Add("App güncelleniyor.");
                    ExtractFile(Application.StartupPath + @"\BilsetNET.7z", textBox1.Text);
                    progressBar1.Value += 10;
                    listBox1.Items.Add("App güncellendi.");
                }
            }
        }
        void VeritabanıEveryoneYetkilendirme()
        {
            listBox1.Items.Add("Veri Tabanları 'Everyone' olarak yetkilendiriliyor....");
            try
            {
                AddFileSecurity(textBox1.Text + @"\BilsetData\Bilset.mdf", @"Everyone", FileSystemRights.FullControl, AccessControlType.Allow);
                AddFileSecurity(textBox1.Text + @"\BilsetData\Bilset_log.ldf", @"Everyone", FileSystemRights.FullControl, AccessControlType.Allow);
                AddFileSecurity(textBox1.Text + @"\BilsetData\BILSET_MASTER.mdf", @"Everyone", FileSystemRights.FullControl, AccessControlType.Allow);
                AddFileSecurity(textBox1.Text + @"\BilsetData\BILSET_MASTER_log.ldf", @"Everyone", FileSystemRights.FullControl, AccessControlType.Allow);
                AddFileSecurity(textBox1.Text + @"\BilsetData\Bilset_EMPTY.mdf", @"Everyone", FileSystemRights.FullControl, AccessControlType.Allow);
                AddFileSecurity(textBox1.Text + @"\BilsetData\Bilset_EMPTY_log.ldf", @"Everyone", FileSystemRights.FullControl, AccessControlType.Allow);
                listBox1.Items.Add("Veri Tabanları 'Everyone' olarak yetkilendirildi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                listBox1.Items.Add(ex);
            }
        }
        void VeriTabanlariAttachEtme()
        {
            listBox1.Items.Add("BILSET Veritabanı oluşturuluyor...");
            database_attach("BILSET", textBox1.Text + @"\BilsetData\Bilset.mdf", textBox1.Text + @"\BilsetData\Bilset_log.ldf");
            listBox1.Items.Add("BILSET oluşturuldu.");

            listBox1.Items.Add("BILSET_MASTER Veritabanı oluşturuluyor...");
            database_attach("BILSET_MASTER", textBox1.Text + @"\BilsetData\BILSET_MASTER.mdf", textBox1.Text + @"\BilsetData\BILSET_MASTER_log.ldf");
            listBox1.Items.Add("BILSET_MASTER oluşturuldu.");
        }
        void KullaniciAtamaVeYetkilendirme()
        {
            listBox1.Items.Add("İlk kullanım için kullanıcı tanımlanıyor....");
            try
            {
                if (SourceConnection.State == System.Data.ConnectionState.Closed)
                {
                    SourceConnection.Open();
                }
                SqlCommand UserAta = new SqlCommand("create login AppUser with password='1234';", SourceConnection);
                UserAta.ExecuteNonQuery();

                SourceConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                listBox1.Items.Add(ex);
                MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }

            listBox1.Items.Add("Kullanıcı Atandı.");

            listBox1.Items.Add("Kullanıcı Yetkilendiriliyor....");
            try
            {
                if (SourceConnection.State == System.Data.ConnectionState.Closed)
                {
                    SourceConnection.Open();
                }
                SqlCommand UserAdmin = new SqlCommand("exec master..sp_addsrvrolemember Bilset,sysadmin;", SourceConnection);
                UserAdmin.ExecuteNonQuery();
                SourceConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kurulum bir hatayla karşılaştı!!\n\n Detay : " + ex.ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                listBox1.Items.Add(ex);
                MessageBox.Show("Kur kapatılacak...", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }
            listBox1.Items.Add("Kullanıcı Yetkilendirildi....");
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            SQLSETUP();

            AppKurulumu();

            VeritabanıEveryoneYetkilendirme();

            VeriTabanlariAttachEtme();

            KullaniciAtamaVeYetkilendirme();

            FireWallnTCPPort();

            CreateShortcut();

            MessageBox.Show("Kurulum Başarıyla Tamamlandı.", "BİLGİ", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Close();
        }
    }
}
