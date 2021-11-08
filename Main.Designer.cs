namespace generatexml
{
    partial class Main
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
            this.ButtonXML = new System.Windows.Forms.Button();
            this.labelVersion = new System.Windows.Forms.Label();
            this.radioButtonBoth = new System.Windows.Forms.RadioButton();
            this.radioButtonXMLOnly = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // ButtonXML
            // 
            this.ButtonXML.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ButtonXML.Location = new System.Drawing.Point(278, 135);
            this.ButtonXML.Name = "ButtonXML";
            this.ButtonXML.Size = new System.Drawing.Size(223, 80);
            this.ButtonXML.TabIndex = 0;
            this.ButtonXML.Text = "Generate XML file(s)\r\nand/or\r\nMS Access database";
            this.ButtonXML.UseVisualStyleBackColor = true;
            this.ButtonXML.Click += new System.EventHandler(this.ButtonXML_Click);
            // 
            // labelVersion
            // 
            this.labelVersion.AutoSize = true;
            this.labelVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelVersion.Location = new System.Drawing.Point(273, 92);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Size = new System.Drawing.Size(99, 25);
            this.labelVersion.TabIndex = 1;
            this.labelVersion.Text = "Version: ";
            // 
            // radioButtonBoth
            // 
            this.radioButtonBoth.AutoSize = true;
            this.radioButtonBoth.Location = new System.Drawing.Point(296, 234);
            this.radioButtonBoth.Name = "radioButtonBoth";
            this.radioButtonBoth.Size = new System.Drawing.Size(163, 17);
            this.radioButtonBoth.TabIndex = 3;
            this.radioButtonBoth.Text = "Both Database and XML files";
            this.radioButtonBoth.UseVisualStyleBackColor = true;
            // 
            // radioButtonXMLOnly
            // 
            this.radioButtonXMLOnly.AutoSize = true;
            this.radioButtonXMLOnly.Location = new System.Drawing.Point(296, 257);
            this.radioButtonXMLOnly.Name = "radioButtonXMLOnly";
            this.radioButtonXMLOnly.Size = new System.Drawing.Size(90, 17);
            this.radioButtonXMLOnly.TabIndex = 4;
            this.radioButtonXMLOnly.Text = "XML files only";
            this.radioButtonXMLOnly.UseVisualStyleBackColor = true;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.radioButtonXMLOnly);
            this.Controls.Add(this.radioButtonBoth);
            this.Controls.Add(this.labelVersion);
            this.Controls.Add(this.ButtonXML);
            this.Name = "Main";
            this.Text = "GiST Config";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ButtonXML;
        private System.Windows.Forms.Label labelVersion;
        private System.Windows.Forms.RadioButton radioButtonBoth;
        private System.Windows.Forms.RadioButton radioButtonXMLOnly;
    }
}

