
namespace aQua_Institut_Druckmanger
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dateienBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.Dateiname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NeuerDateiname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.anzahlKopienDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.umlautEntfernenDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.pDFErzeugenDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.WasserzeichenHinzufügen = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateienBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Dateiname,
            this.NeuerDateiname,
            this.anzahlKopienDataGridViewTextBoxColumn,
            this.umlautEntfernenDataGridViewCheckBoxColumn,
            this.pDFErzeugenDataGridViewCheckBoxColumn,
            this.WasserzeichenHinzufügen});
            this.dataGridView1.DataSource = this.dateienBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(698, 481);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // dateienBindingSource
            // 
            this.dateienBindingSource.DataSource = typeof(aQua_Institut_Druckmanger.Datei);
            this.dateienBindingSource.CurrentChanged += new System.EventHandler(this.dateienBindingSource_CurrentChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(753, 27);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(204, 52);
            this.button1.TabIndex = 1;
            this.button1.Text = "Dokumente auswählen";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(753, 127);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(204, 52);
            this.button2.TabIndex = 2;
            this.button2.Text = "Drucken";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(753, 227);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(204, 52);
            this.button3.TabIndex = 3;
            this.button3.Text = "Speichern unter";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(59, 552);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(582, 56);
            this.button8.TabIndex = 8;
            this.button8.Text = "Liste leeren";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::aQua_Institut_Druckmanger.Properties.Resources.aqualogo;
            this.pictureBox1.Location = new System.Drawing.Point(753, 552);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(204, 87);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(753, 425);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(204, 52);
            this.button5.TabIndex = 11;
            this.button5.Text = "Rückmeldung per Outlook";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click_1);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(753, 326);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(204, 52);
            this.button4.TabIndex = 12;
            this.button4.Text = "Hilfe";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Dateiname
            // 
            this.Dateiname.DataPropertyName = "Dateiname";
            this.Dateiname.HeaderText = "Dateiname";
            this.Dateiname.Name = "Dateiname";
            // 
            // NeuerDateiname
            // 
            this.NeuerDateiname.DataPropertyName = "NeuerDateiname";
            this.NeuerDateiname.HeaderText = "Neuer Dateiname";
            this.NeuerDateiname.Name = "NeuerDateiname";
            // 
            // anzahlKopienDataGridViewTextBoxColumn
            // 
            this.anzahlKopienDataGridViewTextBoxColumn.DataPropertyName = "AnzahlKopien";
            this.anzahlKopienDataGridViewTextBoxColumn.HeaderText = "Anzahl Kopien";
            this.anzahlKopienDataGridViewTextBoxColumn.Name = "anzahlKopienDataGridViewTextBoxColumn";
            // 
            // umlautEntfernenDataGridViewCheckBoxColumn
            // 
            this.umlautEntfernenDataGridViewCheckBoxColumn.DataPropertyName = "UmlautEntfernen";
            this.umlautEntfernenDataGridViewCheckBoxColumn.HeaderText = "Umlaut entfernen";
            this.umlautEntfernenDataGridViewCheckBoxColumn.Name = "umlautEntfernenDataGridViewCheckBoxColumn";
            // 
            // pDFErzeugenDataGridViewCheckBoxColumn
            // 
            this.pDFErzeugenDataGridViewCheckBoxColumn.DataPropertyName = "PDFErzeugen";
            this.pDFErzeugenDataGridViewCheckBoxColumn.HeaderText = "PDF erzeugen";
            this.pDFErzeugenDataGridViewCheckBoxColumn.Name = "pDFErzeugenDataGridViewCheckBoxColumn";
            // 
            // WasserzeichenHinzufügen
            // 
            this.WasserzeichenHinzufügen.DataPropertyName = "WasserzeichenHinzufügen";
            this.WasserzeichenHinzufügen.HeaderText = "WasserzeichenHinzufügen";
            this.WasserzeichenHinzufügen.Name = "WasserzeichenHinzufügen";
            this.WasserzeichenHinzufügen.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.WasserzeichenHinzufügen.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(78)))), ((int)(((byte)(168)))));
            this.ClientSize = new System.Drawing.Size(1001, 689);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "aQua Druckmanager V.1.0.0 (Aktualisiert am 04.05.2022)";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateienBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn neueDateiennamenDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn waasserzeichenHinzufügenDataGridViewCheckBoxColumn;
        private System.Windows.Forms.BindingSource dateienBindingSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn dateinamenDataGridViewTextBoxColumn;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Dateiname;
        private System.Windows.Forms.DataGridViewTextBoxColumn NeuerDateiname;
        private System.Windows.Forms.DataGridViewTextBoxColumn anzahlKopienDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn umlautEntfernenDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn pDFErzeugenDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn WasserzeichenHinzufügen;
    }
}

