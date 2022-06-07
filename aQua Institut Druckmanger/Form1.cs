using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading;
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
                    
                    /* Console.WriteLine(datei.pfad);                //C:\Users\J.Lee\Desktop\Documents\filename.docx
                     Console.WriteLine(datei.Dateiname);             //filename.docx
                     Console.WriteLine(datei.fileDirectory);         //C:\Users\J.Lee\Desktop\Documents
                     Console.WriteLine(datei.newPath);               //C:\Users\J.Lee\Desktop\newselectedfolder\filename
                     Console.WriteLine(datei.NeuerDateiname);        //filename
                     Console.WriteLine(datei.newFileDirectory);      //C:\Users\J.Lee\Desktop\newselectedfolder
                     */
                    if (string.IsNullOrEmpty(datei.NeuerDateiname) == false) 
                    {
                        FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                        fbd1.ShowNewFolderButton = true;
                        if(fbd1.ShowDialog() == DialogResult.OK)
                        {
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
                            datei.fileDirectory = Path.GetDirectoryName(datei.pfad);
                            datei.newPath = Path.GetFullPath(fbd1.SelectedPath + "\\" + datei.NeuerDateiname);
                            SaveAs(datei);
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
        private void AddTextWatermark(Datei datei) 
        {
            object oMissing = Missing.Value;
            Microsoft.Office.Interop.Word.Document doc = null;
            Microsoft.Office.Interop.Word.Application app= null;
            doc = app.Documents.Open(datei.pfad);
            doc.Activate();
            Microsoft.Office.Interop.Word.Shape textWatermark = null;
            foreach(Microsoft.Office.Interop.Word.Section section in doc.Sections)
            {
                textWatermark = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1,
                                datei.TextWasserzeichen, "Arial", (float)60, Microsoft.Office.Core.MsoTriState.msoTrue, 
                                Microsoft.Office.Core.MsoTriState.msoFalse, 0, 0, ref oMissing);
                textWatermark.Select(ref oMissing);
                textWatermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                textWatermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                textWatermark.Fill.Solid();
                textWatermark.Fill.ForeColor.RGB = (Int32)Microsoft.Office.Interop.Word.WdColor.wdColorGray10;
                textWatermark.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                textWatermark.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                textWatermark.Left = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter;
                textWatermark.Top = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter;
                textWatermark.Height = app.InchesToPoints(2.4f);
                textWatermark.Width = app.InchesToPoints(6f);
            }
            doc.SaveAs2(datei.newPath + " mit Wasserzeichen.docx");
            doc.Close();
            app.Quit();
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
            string fileExtension; //.extension
            fileExtension = Path.GetExtension(datei.pfad);
            switch (fileExtension)
            {
                case ".pptx": case ".ppt":
                    Aspose.Slides.Presentation presentation = new Aspose.Slides.Presentation(datei.pfad);
                    presentation.Save(datei.NeuerDateiname + ".pdf", Aspose.Slides.Export.SaveFormat.Pdf);
                    break;
                case ".doc": case".docx":
                    Aspose.Words.Document doc = new Aspose.Words.Document(datei.pfad);
                    doc.Save(datei.newPath + ".pdf");
                    break;
                case ".xlsx":
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(datei.pfad);
                    workbook.Save(datei.NeuerDateiname+".pdf", Aspose.Cells.SaveFormat.Pdf);
                    break;  
                default: MessageBox.Show("Dieser Datei Typ wird nicht unterstützt. Nur pptx, ppt, doc, docx, xlsx Typen sind zu verwenden.");
                    break;
            }
            
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    } 
}

