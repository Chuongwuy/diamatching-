namespace diematching
{
    partial class Diematching
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
            this.browsebutton = new System.Windows.Forms.Button();
            this.datagridview = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.rawtextbox = new System.Windows.Forms.TextBox();
            this.resultlabel = new System.Windows.Forms.Label();
            this.processedtextbox = new System.Windows.Forms.TextBox();
            this.generatebutton = new System.Windows.Forms.Button();
            this.rawlabel = new System.Windows.Forms.Label();
            this.processedlabel = new System.Windows.Forms.Label();
            this.cancelbutton = new System.Windows.Forms.Button();
            this.loadbutton = new System.Windows.Forms.Button();
            this.savelabel = new System.Windows.Forms.Label();
            this.savetextbox = new System.Windows.Forms.TextBox();
            this.savebutton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.datagridview)).BeginInit();
            this.SuspendLayout();
            // 
            // browsebutton
            // 
            this.browsebutton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.browsebutton.Location = new System.Drawing.Point(1159, 88);
            this.browsebutton.Name = "browsebutton";
            this.browsebutton.Size = new System.Drawing.Size(102, 39);
            this.browsebutton.TabIndex = 0;
            this.browsebutton.Text = "Browse";
            this.browsebutton.UseVisualStyleBackColor = false;
            this.browsebutton.Click += new System.EventHandler(this.button1_Click);
            // 
            // datagridview
            // 
            this.datagridview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.datagridview.Location = new System.Drawing.Point(12, 62);
            this.datagridview.Name = "datagridview";
            this.datagridview.RowHeadersWidth = 51;
            this.datagridview.RowTemplate.Height = 24;
            this.datagridview.Size = new System.Drawing.Size(819, 418);
            this.datagridview.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(808, 138);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 16);
            this.label1.TabIndex = 2;
            // 
            // rawtextbox
            // 
            this.rawtextbox.Location = new System.Drawing.Point(867, 96);
            this.rawtextbox.Name = "rawtextbox";
            this.rawtextbox.Size = new System.Drawing.Size(255, 22);
            this.rawtextbox.TabIndex = 3;
            // 
            // resultlabel
            // 
            this.resultlabel.AutoSize = true;
            this.resultlabel.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.resultlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.resultlabel.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.resultlabel.Location = new System.Drawing.Point(12, 501);
            this.resultlabel.Name = "resultlabel";
            this.resultlabel.Size = new System.Drawing.Size(103, 25);
            this.resultlabel.TabIndex = 4;
            this.resultlabel.Text = "RESULT:";
            this.resultlabel.Click += new System.EventHandler(this.resultlabel_Click);
            // 
            // processedtextbox
            // 
            this.processedtextbox.Location = new System.Drawing.Point(869, 335);
            this.processedtextbox.Name = "processedtextbox";
            this.processedtextbox.Size = new System.Drawing.Size(379, 22);
            this.processedtextbox.TabIndex = 5;
            // 
            // generatebutton
            // 
            this.generatebutton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.generatebutton.Location = new System.Drawing.Point(1020, 244);
            this.generatebutton.Name = "generatebutton";
            this.generatebutton.Size = new System.Drawing.Size(102, 39);
            this.generatebutton.TabIndex = 6;
            this.generatebutton.Text = "Generate";
            this.generatebutton.UseVisualStyleBackColor = false;
            // 
            // rawlabel
            // 
            this.rawlabel.AutoSize = true;
            this.rawlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rawlabel.Location = new System.Drawing.Point(863, 61);
            this.rawlabel.Name = "rawlabel";
            this.rawlabel.Size = new System.Drawing.Size(92, 20);
            this.rawlabel.TabIndex = 7;
            this.rawlabel.Text = "Input data";
            // 
            // processedlabel
            // 
            this.processedlabel.AutoSize = true;
            this.processedlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processedlabel.Location = new System.Drawing.Point(865, 301);
            this.processedlabel.Name = "processedlabel";
            this.processedlabel.Size = new System.Drawing.Size(198, 20);
            this.processedlabel.TabIndex = 8;
            this.processedlabel.Text = "Processed output data";
            // 
            // cancelbutton
            // 
            this.cancelbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelbutton.Location = new System.Drawing.Point(1123, 393);
            this.cancelbutton.Name = "cancelbutton";
            this.cancelbutton.Size = new System.Drawing.Size(138, 54);
            this.cancelbutton.TabIndex = 9;
            this.cancelbutton.Text = "Cancel";
            this.cancelbutton.UseVisualStyleBackColor = true;
            // 
            // loadbutton
            // 
            this.loadbutton.BackColor = System.Drawing.SystemColors.Info;
            this.loadbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loadbutton.Location = new System.Drawing.Point(867, 393);
            this.loadbutton.Name = "loadbutton";
            this.loadbutton.Size = new System.Drawing.Size(230, 54);
            this.loadbutton.TabIndex = 10;
            this.loadbutton.Text = "Load Data Table";
            this.loadbutton.UseVisualStyleBackColor = false;
            // 
            // savelabel
            // 
            this.savelabel.AutoSize = true;
            this.savelabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.savelabel.Location = new System.Drawing.Point(863, 139);
            this.savelabel.Name = "savelabel";
            this.savelabel.Size = new System.Drawing.Size(107, 20);
            this.savelabel.TabIndex = 11;
            this.savelabel.Text = "Output data";
            // 
            // savetextbox
            // 
            this.savetextbox.Location = new System.Drawing.Point(869, 180);
            this.savetextbox.Name = "savetextbox";
            this.savetextbox.Size = new System.Drawing.Size(255, 22);
            this.savetextbox.TabIndex = 12;
            // 
            // savebutton
            // 
            this.savebutton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.savebutton.Location = new System.Drawing.Point(1159, 172);
            this.savebutton.Name = "savebutton";
            this.savebutton.Size = new System.Drawing.Size(102, 39);
            this.savebutton.TabIndex = 13;
            this.savebutton.Text = "Browse";
            this.savebutton.UseVisualStyleBackColor = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 19.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(505, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(286, 38);
            this.label2.TabIndex = 14;
            this.label2.Text = "DIA MATCHING ";
            // 
            // Diematching
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1307, 549);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.savebutton);
            this.Controls.Add(this.savetextbox);
            this.Controls.Add(this.savelabel);
            this.Controls.Add(this.loadbutton);
            this.Controls.Add(this.cancelbutton);
            this.Controls.Add(this.processedlabel);
            this.Controls.Add(this.rawlabel);
            this.Controls.Add(this.generatebutton);
            this.Controls.Add(this.processedtextbox);
            this.Controls.Add(this.resultlabel);
            this.Controls.Add(this.rawtextbox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.datagridview);
            this.Controls.Add(this.browsebutton);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Diematching";
            this.Text = "Dia matching app ";
            this.Load += new System.EventHandler(this.Diematching_Load);
            ((System.ComponentModel.ISupportInitialize)(this.datagridview)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button browsebutton;
        private System.Windows.Forms.DataGridView datagridview;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox rawtextbox;
        private System.Windows.Forms.Label resultlabel;
        private System.Windows.Forms.TextBox processedtextbox;
        private System.Windows.Forms.Button generatebutton;
        private System.Windows.Forms.Label rawlabel;
        private System.Windows.Forms.Label processedlabel;
        private System.Windows.Forms.Button cancelbutton;
        private System.Windows.Forms.Button loadbutton;
        private System.Windows.Forms.Label savelabel;
        private System.Windows.Forms.TextBox savetextbox;
        private System.Windows.Forms.Button savebutton;
        private System.Windows.Forms.Label label2;
    }
}

