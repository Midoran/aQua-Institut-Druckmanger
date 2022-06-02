using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using EnvDTE;
using EnvDTE80;
using System.Runtime.InteropServices;
using Document = Microsoft.Office.Interop.Word.Document;
using Application = Microsoft.Office.Interop.Word.Application;
using Font = Microsoft.Office.Interop.Word.Font;
using System.Web;
using RawPrint;

namespace aQua_Institut_Druckmanger //drucken ohne Dialog
{
    public partial class Form1 : Form
    {
        BindingList<Datei> dateienList = new BindingList<Datei>();
        //List<data type> list name = new List<data type> 여기서 data type이 class인 것. 그 class의 이름은 Dateien인 것이고.
        // Dateien이란 클래스들의 목록, 리스트를 클래스로 쓰는 것. 그것의 한 인스턴스가 DateienList인 것.
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
                    string extension;
                    extension = Path.GetExtension(datei.pfad);
                    
                    if (extension == ".docx" || extension == ".doc") //if the file extension is docx or doc
                    {
                        Document doc = word.Documents.Open(datei.pfad);
                        object copies = datei.AnzahlKopien;
                        object pages = "";
                        object range = Microsoft.Office.Interop.Word.WdPrintOutRange.wdPrintAllDocument;
                        object items = Microsoft.Office.Interop.Word.WdPrintOutItem.wdPrintDocumentContent;
                        object pageType = Microsoft.Office.Interop.Word.WdPrintOutPages.wdPrintAllPages;
                        object oTrue = true;
                        object oFalse = false;
                        object missing = Type.Missing;
                        object outputFileName = "";
                        doc.PrintOut(ref oTrue, ref oFalse, ref range, ref outputFileName, ref missing, ref missing,
                                     ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue, ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);
                    }
                    else
                    {
                        printNonWordFiles(datei);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler! Die Dokumente können nicht gedruckt werden: " + ex.Message);
            }
        }
        private void printNonWordFiles(Datei datei)
        {
          
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Datei datei in dateienList)
                {
                    string extension;
                    extension = Path.GetExtension(datei.pfad);
                    /* Console.WriteLine(datei.pfad);                //C:\Users\J.Lee\Desktop\Documents\filename.docx
                     Console.WriteLine(datei.Dateiname);             //filename.docx
                     Console.WriteLine(datei.fileDirectory);         //C:\Users\J.Lee\Desktop\Documents
                     Console.WriteLine(datei.newPath);               //C:\Users\J.Lee\Desktop\newselectedfolder\filename
                     Console.WriteLine(datei.NeuerDateiname);        //filename
                     Console.WriteLine(datei.newFileDirectory);      //C:\Users\J.Lee\Desktop\newselectedfolder
                     */
                    if (string.IsNullOrEmpty(datei.WasserzeichenHinzufügen) == false) //watermark
                    {
                        FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                        fbd1.ShowNewFolderButton = true;
                        if (fbd1.ShowDialog() == DialogResult.OK)
                        {
                            datei.fileDirectory = Path.GetDirectoryName(datei.pfad);
                            datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                            datei.newFileDirectory = Path.GetDirectoryName(datei.newPath);
                            AddWatermark(datei);
                        }
                    }
                    if (datei.UmlautEntfernen == true)//umlaut
                    {
                        FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                        fbd1.ShowNewFolderButton = true;
                        if (fbd1.ShowDialog() == DialogResult.OK)
                        {
                            datei.fileDirectory = Path.GetDirectoryName(datei.pfad);
                            datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                            datei.newFileDirectory = Path.GetDirectoryName(datei.newPath);
                            UmlautEntfernen(datei);
                        }
                    }
                    if (datei.PDFErzeugen == true)//PDF
                    {
                        FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                        fbd1.ShowNewFolderButton = true;
                        if (fbd1.ShowDialog() == DialogResult.OK)
                        {
                            datei.fileDirectory = Path.GetDirectoryName(datei.pfad);
                            datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                            datei.newFileDirectory = Path.GetDirectoryName(datei.newPath);
                            PdfErzeugen(datei);
                        }
                    }
                    if (string.IsNullOrEmpty(datei.NeuerDateiname) == false) 
                    {
                        FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                        fbd1.ShowNewFolderButton = true;
                        if (fbd1.ShowDialog() == DialogResult.OK&&extension==".doc"&&extension==".docx")
                        {
                            datei.fileDirectory = Path.GetDirectoryName(datei.pfad);
                            datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                            datei.newFileDirectory = Path.GetDirectoryName(datei.newPath);
                            SaveAs(datei);
                        }
                        else
                        {
                            MessageBox.Show("NUR Word Dateien können umgewandelt werden. Andere Dateiformaten können nur zu PDF umgewandelt werden.");
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
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
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

        private void AddWatermark(Datei datei) 
        {
            object o = Missing.Value;
            object oT = true;
            object oF = false;
            //Microsoft.Office.Interop.Word.Application oword = new Word
            /*Aspose.Words.Document doc = new Aspose.Words.Document(datei.pfad);
            Aspose.Words.TextWatermarkOptions options = new Aspose.Words.TextWatermarkOptions()
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.LightGray,
                Layout = Aspose.Words.WatermarkLayout.Diagonal,
                IsSemitrasparent = true
            };
            doc.Watermark.SetText(datei.WasserzeichenHinzufügen, options);
            doc.Save(datei.newPath + " mit Wasserzeichen.docx");*/
            /* object o = Missing.Value;
             object oT = true;
             object oF = false;
             Microsoft.Office.Interop.Word.Application oWord = new Application();
             Microsoft.Office.Interop.Word.Document oWordDoc = new Document();
             oWordDoc = oWord.Documents.Open(datei.pfad, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
             oWordDoc.Activate();
             oWordDoc = oWord.Documents.Add(ref o, ref o, ref o, ref o);
             Microsoft.Office.Interop.Word.Shape logoWatermark = oWord.Selection.HeaderFooter.Shapes.AddTextEffect(
                 Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1,
                 datei.WasserzeichenHinzufügen, "Arial", (float)60,
                 Microsoft.Office.Core.MsoTriState.msoTrue,
                 Microsoft.Office.Core.MsoTriState.msoFalse, 
                 0, 0, ref o);
             logoWatermark.Select(ref o);
             logoWatermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
             logoWatermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
             logoWatermark.Fill.Solid();
             logoWatermark.Fill.ForeColor.RGB = (Int32)Microsoft.Office.Interop.Word.WdColor.wdColorGray30;
             logoWatermark.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
             logoWatermark.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
             logoWatermark.Left = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter;
             logoWatermark.Top = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter;
             logoWatermark.Height = oWord.InchesToPoints(2.4f);
             logoWatermark.Width = oWord.InchesToPoints(6f);
             oWord.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;

             oWordDoc.SaveAs2(datei.newPath + " mit Wasserzeichen.docx", ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
             oWordDoc.Close();
             oWord.Quit();*/
        }
       
    
        private void UmlautEntfernen(Datei datei)
        {
            Microsoft.Office.Interop.Word.Application app = new Application();
            object o = Missing.Value;
            object t = true;
            object f = false;
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
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application app, object findText, object replaceWithText)
        {
            object oFalse = false;
            object oTrue = true;
            object o = Missing.Value;
             
            app.Selection.Find.Execute(ref findText, oFalse, oFalse, oFalse, oFalse, oFalse, oTrue, oFalse, oFalse, ref replaceWithText, 2, oFalse, oFalse, oFalse, oFalse);
        }
        private void PdfErzeugen(Datei datei)
        {
            ProcessStartInfo pdfProcessInfo = new ProcessStartInfo()
            {
                Verb = "print",
                CreateNoWindow = true,
                FileName = datei.pfad,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            System.Diagnostics.Process pdfProcess = new System.Diagnostics.Process();
            pdfProcess.StartInfo = pdfProcessInfo;
            pdfProcess.Start();
            pdfProcess.WaitForInputIdle();
        }
    } 
}

