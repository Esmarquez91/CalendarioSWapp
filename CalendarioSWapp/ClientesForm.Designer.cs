namespace CalendarioSWapp
{
    partial class ClientesForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ClientesForm));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.DGClientes = new System.Windows.Forms.DataGridView();
            this.TxtBoxFiltro = new System.Windows.Forms.TextBox();
            this.BtnFiltrar = new System.Windows.Forms.Button();
            this.TicketSelected = new System.Windows.Forms.TextBox();
            this.LabelResultados = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGClientes)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 5;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Controls.Add(this.DGClientes, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.TxtBoxFiltro, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.BtnFiltrar, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.TicketSelected, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.LabelResultados, 3, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(12, 12);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 10;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 6.845942F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 6.845942F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78852F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78852F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78851F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78851F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78851F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78851F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78851F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.78851F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(200, 426);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // DGClientes
            // 
            this.DGClientes.AllowUserToAddRows = false;
            this.DGClientes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DGClientes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGClientes.BackgroundColor = System.Drawing.SystemColors.Control;
            this.DGClientes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableLayoutPanel1.SetColumnSpan(this.DGClientes, 5);
            this.DGClientes.Location = new System.Drawing.Point(3, 61);
            this.DGClientes.Name = "DGClientes";
            this.DGClientes.RowHeadersVisible = false;
            this.tableLayoutPanel1.SetRowSpan(this.DGClientes, 8);
            this.DGClientes.Size = new System.Drawing.Size(194, 362);
            this.DGClientes.TabIndex = 0;
            this.DGClientes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DGClientes_KeyDown);
            // 
            // TxtBoxFiltro
            // 
            this.TxtBoxFiltro.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.SetColumnSpan(this.TxtBoxFiltro, 3);
            this.TxtBoxFiltro.Location = new System.Drawing.Point(3, 35);
            this.TxtBoxFiltro.Name = "TxtBoxFiltro";
            this.TxtBoxFiltro.Size = new System.Drawing.Size(94, 20);
            this.TxtBoxFiltro.TabIndex = 1;
            this.TxtBoxFiltro.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtBoxFiltro_KeyDown);
            // 
            // BtnFiltrar
            // 
            this.BtnFiltrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tableLayoutPanel1.SetColumnSpan(this.BtnFiltrar, 2);
            this.BtnFiltrar.Location = new System.Drawing.Point(103, 32);
            this.BtnFiltrar.Name = "BtnFiltrar";
            this.BtnFiltrar.Size = new System.Drawing.Size(94, 23);
            this.BtnFiltrar.TabIndex = 2;
            this.BtnFiltrar.Text = "Filtrar CRT";
            this.BtnFiltrar.UseVisualStyleBackColor = true;
            this.BtnFiltrar.Click += new System.EventHandler(this.BtnFiltrar_Click);
            this.BtnFiltrar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.BtnFiltrar_KeyDown);
            // 
            // TicketSelected
            // 
            this.TicketSelected.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.TicketSelected.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tableLayoutPanel1.SetColumnSpan(this.TicketSelected, 3);
            this.TicketSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TicketSelected.Location = new System.Drawing.Point(3, 10);
            this.TicketSelected.Name = "TicketSelected";
            this.TicketSelected.ReadOnly = true;
            this.TicketSelected.Size = new System.Drawing.Size(94, 16);
            this.TicketSelected.TabIndex = 3;
            this.TicketSelected.Text = "SWX";
            // 
            // LabelResultados
            // 
            this.LabelResultados.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.LabelResultados.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.LabelResultados, 2);
            this.LabelResultados.Location = new System.Drawing.Point(103, 8);
            this.LabelResultados.Name = "LabelResultados";
            this.LabelResultados.Size = new System.Drawing.Size(66, 13);
            this.LabelResultados.TabIndex = 4;
            this.LabelResultados.Text = "Resultados: ";
            // 
            // ClientesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(228, 450);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ClientesForm";
            this.Text = "ClientesForm";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ClientesForm_KeyDown);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGClientes)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView DGClientes;
        private System.Windows.Forms.TextBox TxtBoxFiltro;
        private System.Windows.Forms.Button BtnFiltrar;
        private System.Windows.Forms.TextBox TicketSelected;
        private System.Windows.Forms.Label LabelResultados;
    }
}