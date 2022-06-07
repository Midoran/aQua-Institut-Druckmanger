using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace aQua_Institut_Druckmanger
{
    public class Datei  
    {
        [DisplayName("Dateiname")] 
        public string Dateiname { get; set; }
        [DisplayName("Neuer Dateiname")]
        public string NeuerDateiname { get; set; }
        [DisplayName("Anzahl Kopien")]
        public int  AnzahlKopien { get; set; } = 1;
        
        [DisplayName("Umlaut entfernen")]
        public bool UmlautEntfernen { get; set; }
        [DisplayName("PDF")]
        public bool PDFErzeugen { get; set; }
        [DisplayName("Text Wasserzeichen")]
        public string TextWasserzeichen { get; set; }
        public string pfad { get; set; }
        public string newFileDirectory { get; set; }
        public string fileDirectory { get; set; }
        public string newPath { get; set; }
        public object filestream { get; set; }
        public Datei()   //Every time this class is created, these lines of code get executed
        {
           
        }
    }

}

