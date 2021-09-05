namespace vl_tools
{
    partial class DrawCatenaryForm
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
            this.groupBoxAll = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.comboBoxRelativeLoadUnit = new System.Windows.Forms.ComboBox();
            this.comboBoxStressUnit = new System.Windows.Forms.ComboBox();
            this.textBoxRelativeLoad = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxStress = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.comboBoxLoadUnit = new System.Windows.Forms.ComboBox();
            this.comboBoxTensionUnit = new System.Windows.Forms.ComboBox();
            this.textBoxLoad = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxTension = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxFm = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.buttonOK = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBoxHScale = new System.Windows.Forms.ComboBox();
            this.comboBoxVScale = new System.Windows.Forms.ComboBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.textBoxGabarit = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.groupBoxAll.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBoxAll
            // 
            this.groupBoxAll.Controls.Add(this.groupBox3);
            this.groupBoxAll.Controls.Add(this.groupBox2);
            this.groupBoxAll.Controls.Add(this.radioButton1);
            this.groupBoxAll.Controls.Add(this.groupBox1);
            this.groupBoxAll.Controls.Add(this.radioButton3);
            this.groupBoxAll.Controls.Add(this.radioButton2);
            this.groupBoxAll.Location = new System.Drawing.Point(12, 12);
            this.groupBoxAll.Name = "groupBoxAll";
            this.groupBoxAll.Size = new System.Drawing.Size(369, 321);
            this.groupBoxAll.TabIndex = 1;
            this.groupBoxAll.TabStop = false;
            this.groupBoxAll.Text = "Укажите способ ввода";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.comboBoxRelativeLoadUnit);
            this.groupBox3.Controls.Add(this.comboBoxStressUnit);
            this.groupBox3.Controls.Add(this.textBoxRelativeLoad);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.textBoxStress);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(19, 231);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(336, 79);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            // 
            // comboBoxRelativeLoadUnit
            // 
            this.comboBoxRelativeLoadUnit.FormattingEnabled = true;
            this.comboBoxRelativeLoadUnit.Items.AddRange(new object[] {
            "Н/(м*мм2) ",
            "кг/(м*мм2) ",
            "даН/(м*мм2) "});
            this.comboBoxRelativeLoadUnit.Location = new System.Drawing.Point(263, 45);
            this.comboBoxRelativeLoadUnit.Name = "comboBoxRelativeLoadUnit";
            this.comboBoxRelativeLoadUnit.Size = new System.Drawing.Size(67, 21);
            this.comboBoxRelativeLoadUnit.TabIndex = 8;
            // 
            // comboBoxStressUnit
            // 
            this.comboBoxStressUnit.FormattingEnabled = true;
            this.comboBoxStressUnit.Items.AddRange(new object[] {
            "Н/мм2 - МПа",
            "кг/мм2",
            "даН/мм2"});
            this.comboBoxStressUnit.Location = new System.Drawing.Point(263, 19);
            this.comboBoxStressUnit.Name = "comboBoxStressUnit";
            this.comboBoxStressUnit.Size = new System.Drawing.Size(67, 21);
            this.comboBoxStressUnit.TabIndex = 7;
            // 
            // textBoxRelativeLoad
            // 
            this.textBoxRelativeLoad.Location = new System.Drawing.Point(157, 46);
            this.textBoxRelativeLoad.Name = "textBoxRelativeLoad";
            this.textBoxRelativeLoad.Size = new System.Drawing.Size(100, 20);
            this.textBoxRelativeLoad.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 48);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(145, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Удельная нагрузка γ=p/S=";
            // 
            // textBoxStress
            // 
            this.textBoxStress.Location = new System.Drawing.Point(157, 20);
            this.textBoxStress.Name = "textBoxStress";
            this.textBoxStress.Size = new System.Drawing.Size(100, 20);
            this.textBoxStress.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Напряжение σ=H/S=";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.comboBoxLoadUnit);
            this.groupBox2.Controls.Add(this.comboBoxTensionUnit);
            this.groupBox2.Controls.Add(this.textBoxLoad);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.textBoxTension);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(19, 118);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(336, 78);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            // 
            // comboBoxLoadUnit
            // 
            this.comboBoxLoadUnit.FormattingEnabled = true;
            this.comboBoxLoadUnit.Items.AddRange(new object[] {
            "Н/м",
            "кг/м",
            "даН/м"});
            this.comboBoxLoadUnit.Location = new System.Drawing.Point(263, 44);
            this.comboBoxLoadUnit.Name = "comboBoxLoadUnit";
            this.comboBoxLoadUnit.Size = new System.Drawing.Size(67, 21);
            this.comboBoxLoadUnit.TabIndex = 7;
            // 
            // comboBoxTensionUnit
            // 
            this.comboBoxTensionUnit.FormattingEnabled = true;
            this.comboBoxTensionUnit.Items.AddRange(new object[] {
            "Н",
            "кг",
            "даН",
            "т"});
            this.comboBoxTensionUnit.Location = new System.Drawing.Point(263, 18);
            this.comboBoxTensionUnit.Name = "comboBoxTensionUnit";
            this.comboBoxTensionUnit.Size = new System.Drawing.Size(67, 21);
            this.comboBoxTensionUnit.TabIndex = 6;
            // 
            // textBoxLoad
            // 
            this.textBoxLoad.Location = new System.Drawing.Point(157, 45);
            this.textBoxLoad.Name = "textBoxLoad";
            this.textBoxLoad.Size = new System.Drawing.Size(100, 20);
            this.textBoxLoad.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(120, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Погонная нагрузка p=";
            // 
            // textBoxTension
            // 
            this.textBoxTension.Location = new System.Drawing.Point(157, 19);
            this.textBoxTension.Name = "textBoxTension";
            this.textBoxTension.Size = new System.Drawing.Size(100, 20);
            this.textBoxTension.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Тяжение H=";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(19, 26);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(106, 17);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Стрела провеса";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxFm);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(19, 42);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(336, 42);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            // 
            // textBoxFm
            // 
            this.textBoxFm.Location = new System.Drawing.Point(188, 13);
            this.textBoxFm.Name = "textBoxFm";
            this.textBoxFm.Size = new System.Drawing.Size(100, 20);
            this.textBoxFm.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(159, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "Стрела провеса \r\nв середине пролета, fmax, м =";
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(19, 213);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(197, 17);
            this.radioButton3.TabIndex = 3;
            this.radioButton3.Text = "Напряжение и удельная нагрузка";
            this.radioButton3.UseVisualStyleBackColor = true;
            this.radioButton3.CheckedChanged += new System.EventHandler(this.radioButton3_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(19, 102);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(178, 17);
            this.radioButton2.TabIndex = 2;
            this.radioButton2.Text = "Тяжение и погонная нагрузка";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(130, 430);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(98, 30);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(234, 430);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(98, 30);
            this.button2.TabIndex = 3;
            this.button2.Text = "Отмена";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(37, 341);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(140, 13);
            this.label6.TabIndex = 4;
            this.label6.Text = "Масштаб горизонтальный";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(37, 367);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(129, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Масштаб вертикальный";
            // 
            // comboBoxHScale
            // 
            this.comboBoxHScale.FormattingEnabled = true;
            this.comboBoxHScale.Items.AddRange(new object[] {
            "1:100",
            "1:200",
            "1:500",
            "1:1000",
            "1:2000",
            "1:5000"});
            this.comboBoxHScale.Location = new System.Drawing.Point(183, 338);
            this.comboBoxHScale.Name = "comboBoxHScale";
            this.comboBoxHScale.Size = new System.Drawing.Size(67, 21);
            this.comboBoxHScale.TabIndex = 8;
            // 
            // comboBoxVScale
            // 
            this.comboBoxVScale.FormattingEnabled = true;
            this.comboBoxVScale.Items.AddRange(new object[] {
            "1:100",
            "1:200",
            "1:500",
            "1:1000",
            "1:2000",
            "1:5000"});
            this.comboBoxVScale.Location = new System.Drawing.Point(183, 365);
            this.comboBoxVScale.Name = "comboBoxVScale";
            this.comboBoxVScale.Size = new System.Drawing.Size(67, 21);
            this.comboBoxVScale.TabIndex = 9;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::vl_tools.Properties.Resources.провис;
            this.pictureBox1.Location = new System.Drawing.Point(387, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(401, 423);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 10;
            this.pictureBox1.TabStop = false;
            // 
            // textBoxGabarit
            // 
            this.textBoxGabarit.Location = new System.Drawing.Point(183, 392);
            this.textBoxGabarit.Name = "textBoxGabarit";
            this.textBoxGabarit.Size = new System.Drawing.Size(67, 20);
            this.textBoxGabarit.TabIndex = 11;
            this.textBoxGabarit.Text = "7";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(37, 395);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(62, 13);
            this.label8.TabIndex = 12;
            this.label8.Text = "Габарит, м";
            // 
            // DrawCatenaryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 472);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.textBoxGabarit);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.comboBoxVScale);
            this.Controls.Add(this.comboBoxHScale);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.groupBoxAll);
            this.Name = "DrawCatenaryForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "DrawCatenaryForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.DrawCatenaryForm_FormClosed);
            this.groupBoxAll.ResumeLayout(false);
            this.groupBoxAll.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.GroupBox groupBoxAll;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.TextBox textBoxRelativeLoad;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxStress;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxLoad;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxTension;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxFm;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxRelativeLoadUnit;
        private System.Windows.Forms.ComboBox comboBoxStressUnit;
        private System.Windows.Forms.ComboBox comboBoxLoadUnit;
        private System.Windows.Forms.ComboBox comboBoxTensionUnit;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBoxHScale;
        private System.Windows.Forms.ComboBox comboBoxVScale;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textBoxGabarit;
        private System.Windows.Forms.Label label8;
    }
}