namespace R2Tech_Valida_Arquivos
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtCaminhoDiretorio = new System.Windows.Forms.TextBox();
            this.btnProcurar = new System.Windows.Forms.Button();
            this.ofd1 = new System.Windows.Forms.OpenFileDialog();
            this.fbd1 = new System.Windows.Forms.FolderBrowserDialog();
            this.ListaLog = new System.Windows.Forms.ListBox();
            this.Descompactar = new System.Windows.Forms.Button();
            this.loader = new System.Windows.Forms.ProgressBar();
            this.btnProcessar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(256, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Informe o caminho do diretório com os PDF\'s e XLS\'s";
            // 
            // txtCaminhoDiretorio
            // 
            this.txtCaminhoDiretorio.Location = new System.Drawing.Point(15, 25);
            this.txtCaminhoDiretorio.Name = "txtCaminhoDiretorio";
            this.txtCaminhoDiretorio.Size = new System.Drawing.Size(503, 20);
            this.txtCaminhoDiretorio.TabIndex = 1;
            // 
            // btnProcurar
            // 
            this.btnProcurar.Location = new System.Drawing.Point(524, 9);
            this.btnProcurar.Name = "btnProcurar";
            this.btnProcurar.Size = new System.Drawing.Size(102, 23);
            this.btnProcurar.TabIndex = 2;
            this.btnProcurar.Text = "Procurar";
            this.btnProcurar.UseVisualStyleBackColor = true;
            this.btnProcurar.Click += new System.EventHandler(this.btnProcurar_Click);
            // 
            // ofd1
            // 
            this.ofd1.FileName = "openFileDialog1";
            // 
            // fbd1
            // 
            this.fbd1.HelpRequest += new System.EventHandler(this.folderBrowserDialog1_HelpRequest);
            // 
            // ListaLog
            // 
            this.ListaLog.FormattingEnabled = true;
            this.ListaLog.Location = new System.Drawing.Point(15, 78);
            this.ListaLog.Name = "ListaLog";
            this.ListaLog.Size = new System.Drawing.Size(720, 329);
            this.ListaLog.TabIndex = 4;
            this.ListaLog.SelectedIndexChanged += new System.EventHandler(this.ListaLog_SelectedIndexChanged);
            // 
            // Descompactar
            // 
            this.Descompactar.Location = new System.Drawing.Point(633, 9);
            this.Descompactar.Name = "Descompactar";
            this.Descompactar.Size = new System.Drawing.Size(102, 23);
            this.Descompactar.TabIndex = 5;
            this.Descompactar.Text = "Descompactar";
            this.Descompactar.UseVisualStyleBackColor = true;
            this.Descompactar.Click += new System.EventHandler(this.Descompactar_Click);
            // 
            // loader
            // 
            this.loader.Location = new System.Drawing.Point(187, 152);
            this.loader.Name = "loader";
            this.loader.Size = new System.Drawing.Size(375, 23);
            this.loader.TabIndex = 6;
            this.loader.Visible = false;
            // 
            // btnProcessar
            // 
            this.btnProcessar.Location = new System.Drawing.Point(633, 38);
            this.btnProcessar.Name = "btnProcessar";
            this.btnProcessar.Size = new System.Drawing.Size(102, 23);
            this.btnProcessar.TabIndex = 9;
            this.btnProcessar.Text = "Processar";
            this.btnProcessar.UseVisualStyleBackColor = true;
            this.btnProcessar.Click += new System.EventHandler(this.btnProcessar_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(747, 415);
            this.Controls.Add(this.btnProcessar);
            this.Controls.Add(this.loader);
            this.Controls.Add(this.Descompactar);
            this.Controls.Add(this.ListaLog);
            this.Controls.Add(this.btnProcurar);
            this.Controls.Add(this.txtCaminhoDiretorio);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "R2Tech Valida Arquivos";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCaminhoDiretorio;
        private System.Windows.Forms.Button btnProcurar;
        private System.Windows.Forms.OpenFileDialog ofd1;
        private System.Windows.Forms.FolderBrowserDialog fbd1;
        private System.Windows.Forms.ListBox ListaLog;
        private System.Windows.Forms.Button Descompactar;
        private System.Windows.Forms.ProgressBar loader;
        private System.Windows.Forms.Button btnProcessar;
    }
}

