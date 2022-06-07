using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;
using Exception = System.Exception;

namespace aQua_Institut_Druckmanger //drucken ohne Dialog
{
    public partial class Form1 : Form
    {
        BindingList<Datei> dateienList = new BindingList<Datei>();
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridView1.DataSource = dateienList;
        }
        private void button1_Click(object sender, EventArgs e)//Dokumente auswählen
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Word(*.docx)| *.docx|PPT(*.pptx)|*.pptx|PDF(*.pdf)|*.pdf|Alle Dateien(*.*)|*.*";
            ofd.Multiselect = true;
            try
            {
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (String path in ofd.FileNames) 
                    {                                     
                        Datei datei = new Datei();  
                        datei.pfad = path; 
                        datei.Dateiname = Path.GetFileName(path);
                        dateienList.Add(datei);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler! Die Datei kann nicht gelesen werden: " + ex.Message);
            }
        }
        private Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application { Visible = false };
        private void button2_Click(object sender, EventArgs e)//Drucken
        {
            try
            {
                foreach (Datei datei in dateienList)
                {
                    
                        ProcessStartInfo ProcessInfo = new ProcessStartInfo()
                        {
                            Verb = "print",
                            CreateNoWindow = true,
                            FileName = datei.pfad,
                            WindowStyle = ProcessWindowStyle.Hidden
                        };
                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        process.StartInfo = ProcessInfo;
                        process.Start();
                        process.WaitForInputIdle();
                        System.Threading.Thread.Sleep(3000);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler! Die Dokumente können nicht gedruckt werden: " + ex.Message);
            }
        }
        private void button3_Click(object sender, EventArgs e)//speichern unter
        {
            try
            {
                foreach (Datei datei in dateienList)
                {

                    /* 
                     Console.WriteLine(datei.Dateiname);             //filename.docx
                     Console.WriteLine(datei.fileDirectory);         //C:\Users\J.Lee\Desktop\Documents
                     Console.WriteLine(datei.pfad);                  //C:\Users\J.Lee\Desktop\Documents\filename.docx
                     Console.WriteLine(datei.newPath);               //C:\Users\J.Lee\Desktop\newselectedfolder\filename
                     Console.WriteLine(datei.newFileDirectory);      //C:\Users\J.Lee\Desktop\newselectedfolder
                     Console.WriteLine(datei.NeuerDateiname);        //filename
                     */
                    if (string.IsNullOrEmpty(datei.NeuerDateiname) == false) 
                    {
                        FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                        fbd1.ShowNewFolderButton = true;
                        datei.fileDirectory = Path.GetDirectoryName(datei.pfad);
                        if (fbd1.ShowDialog() == DialogResult.OK)
                        {
                            datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);

                            if (string.IsNullOrEmpty(datei.TextWasserzeichen) == false) //watermark
                            {
                                //datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                                AddTextWatermark(datei);
                            }
                            if (datei.UmlautEntfernen == true)//umlaut
                            {
                                datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                                UmlautEntfernen(datei);
                            }
                            if (datei.PDFErzeugen == true)//PDF
                            {
                                PdfErzeugen(datei);
                            }
                            else
                            {
                                SaveAs(datei);// file save 
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler! " + ex.Message);
            }
        }
        private void button8_Click(object sender, EventArgs e)//List leeren
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            dateienList.Clear();
        }
        private void pictureBox1_Click(object sender, EventArgs e) //aQua Logo
        {
            System.Diagnostics.Process.Start("https://www.aqua-institut.de");
        }
        private void dateienBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e) //Hilfe
        {
            MessageBox.Show("1. Dateien müssen ausgewählt werden." + Environment.NewLine + Environment.NewLine +
                "2. Nach der Auswahl Dateien können die Dateien auf der Liste gedruckt werden. Anzahl Kopien kann auch eingegeben vor dem Drucken."
                + Environment.NewLine + Environment.NewLine + "3. Nach der Eingabe der Optionen auf Data Grid View kann 'Speichern unter' geclickt werden. Dann werden die Dateien mit eingegebenen Optionen im ausgewählten Ordner gespeichert.");
        }
        private void button5_Click_1(object sender, EventArgs e) //Rückmeldung
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application objApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem mail = null;
                mail = (Microsoft.Office.Interop.Outlook.MailItem)objApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mail.To = "jeongmin.lee@aqua-institut.de";
                mail.Display();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler!: " + ex);
            }
        }
        private void SaveAs(Datei datei)
        {
            File.Copy(Path.Combine(datei.fileDirectory, datei.Dateiname), datei.newPath + ".docx");
        }
        private void AddTextWatermark(Datei datei) //Wasserzeichen
        {
            Microsoft.Office.Interop.Word.Application app = new Application();
            object o = Missing.Value;
            object oTrue = true;
            object oFalse = false;
            Microsoft.Office.Interop.Word.Document doc = null;
            doc = app.Documents.Open(datei.pfad, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
            doc.Activate();
            //add function
            doc.SaveAs2(datei.newPath + " mit Wasserzeichen in Text.docx", ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
            doc.Close();
            app.Quit();
        }
        private void PdfErzeugen(Datei datei)//PDF Erzeugen
        {
            string fileExtension; //.extension
            fileExtension = Path.GetExtension(datei.pfad);
            switch (fileExtension)
            {
                case ".pptx":
                case ".ppt":
                    Aspose.Slides.Presentation presentation = new Aspose.Slides.Presentation(datei.pfad);
                    presentation.Save(datei.newPath + ".pdf", Aspose.Slides.Export.SaveFormat.Pdf);
                    break;
                case ".doc":
                case ".docx":
                    Aspose.Words.Document doc = new Aspose.Words.Document(datei.pfad);
                    doc.Save(datei.newPath + ".pdf");
                    break;
                case ".xlsx":
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(datei.pfad);
                    workbook.Save(datei.newPath + ".pdf");
                    break;
                default:
                    MessageBox.Show("Dieser Datei Typ wird nicht unterstützt. Nur pptx, ppt, doc, docx, xlsx Typen sind zu verwenden.");
                    break;
            }
        }
        private void UmlautEntfernen(Datei datei)//Umlaut entfernen
        {
            Microsoft.Office.Interop.Word.Application app = new Application();
            object o = Missing.Value;
            object oTrue = true;
            object oFalse = false;
            Microsoft.Office.Interop.Word.Document doc = null;
            doc = app.Documents.Open(datei.pfad, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
            doc.Activate();
            this.FindAndReplace(app, "ö", "oe");
            this.FindAndReplace(app, "Ö", "Oe");
            this.FindAndReplace(app, "ä", "ae");
            this.FindAndReplace(app, "Ä", "Ae");
            this.FindAndReplace(app, "ü", "ue");
            this.FindAndReplace(app, "Ü", "Üe");
            doc.SaveAs2(datei.newPath+" ohne Umlaute.docx", ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
            doc.Close();
            app.Quit();
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application app, object findText, object replaceWithText)//Umlaut entfernen
        {
            object oFalse = false;
            object oTrue = true;
            object o = Missing.Value;
             
            app.Selection.Find.Execute(ref findText, oFalse, oFalse, oFalse, oFalse, oFalse, oTrue, oFalse, oFalse, ref replaceWithText, 2, oFalse, oFalse, oFalse, oFalse);
        }
        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    } 
}

