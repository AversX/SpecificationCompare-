namespace SpecificationCompare
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.oldVerLabel = new MetroFramework.Controls.MetroLabel();
            this.SrchSpecOldBtn = new MetroFramework.Controls.MetroButton();
            this.oldVerTxt = new MetroFramework.Controls.MetroTextBox();
            this.newVerTxt = new MetroFramework.Controls.MetroTextBox();
            this.SrchSpecNewBtn = new MetroFramework.Controls.MetroButton();
            this.compareBtn = new MetroFramework.Controls.MetroButton();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.metroLabel2 = new MetroFramework.Controls.MetroLabel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.metroLabel3 = new MetroFramework.Controls.MetroLabel();
            this.metroLabel4 = new MetroFramework.Controls.MetroLabel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.metroLabel5 = new MetroFramework.Controls.MetroLabel();
            this.artColTBox = new MetroFramework.Controls.MetroTextBox();
            this.headerChBox = new MetroFramework.Controls.MetroCheckBox();
            this.nameColTBox = new MetroFramework.Controls.MetroTextBox();
            this.metroLabel6 = new MetroFramework.Controls.MetroLabel();
            this.SuspendLayout();
            // 
            // oldVerLabel
            // 
            this.oldVerLabel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.oldVerLabel.AutoSize = true;
            this.oldVerLabel.Location = new System.Drawing.Point(5, 182);
            this.oldVerLabel.Name = "oldVerLabel";
            this.oldVerLabel.Size = new System.Drawing.Size(259, 19);
            this.oldVerLabel.Style = MetroFramework.MetroColorStyle.Yellow;
            this.oldVerLabel.TabIndex = 1;
            this.oldVerLabel.Text = "Выберите спецификации для сравнения";
            this.oldVerLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.oldVerLabel.UseStyleColors = true;
            // 
            // SrchSpecOldBtn
            // 
            this.SrchSpecOldBtn.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.SrchSpecOldBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SrchSpecOldBtn.Location = new System.Drawing.Point(334, 247);
            this.SrchSpecOldBtn.Name = "SrchSpecOldBtn";
            this.SrchSpecOldBtn.Size = new System.Drawing.Size(59, 23);
            this.SrchSpecOldBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.SrchSpecOldBtn.TabIndex = 2;
            this.SrchSpecOldBtn.TabStop = false;
            this.SrchSpecOldBtn.Text = "Обзор";
            this.SrchSpecOldBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.SrchSpecOldBtn.Click += new System.EventHandler(this.SrchSpecOldBtn_Click);
            // 
            // oldVerTxt
            // 
            this.oldVerTxt.AllowDrop = true;
            this.oldVerTxt.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.oldVerTxt.Location = new System.Drawing.Point(5, 218);
            this.oldVerTxt.Name = "oldVerTxt";
            this.oldVerTxt.PromptText = "Текущая версия";
            this.oldVerTxt.Size = new System.Drawing.Size(388, 23);
            this.oldVerTxt.TabIndex = 3;
            this.oldVerTxt.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.oldVerTxt.DragDrop += new System.Windows.Forms.DragEventHandler(this.OldVerTxt_DragDrop);
            this.oldVerTxt.DragEnter += new System.Windows.Forms.DragEventHandler(this.OldVerTxt_DragEnter);
            // 
            // newVerTxt
            // 
            this.newVerTxt.AllowDrop = true;
            this.newVerTxt.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.newVerTxt.Location = new System.Drawing.Point(5, 290);
            this.newVerTxt.Name = "newVerTxt";
            this.newVerTxt.PromptText = "Новая версия";
            this.newVerTxt.Size = new System.Drawing.Size(388, 23);
            this.newVerTxt.TabIndex = 5;
            this.newVerTxt.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.newVerTxt.DragDrop += new System.Windows.Forms.DragEventHandler(this.NewVerTxt_DragDrop);
            this.newVerTxt.DragEnter += new System.Windows.Forms.DragEventHandler(this.NewVerTxt_DragEnter);
            // 
            // SrchSpecNewBtn
            // 
            this.SrchSpecNewBtn.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.SrchSpecNewBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SrchSpecNewBtn.Location = new System.Drawing.Point(334, 319);
            this.SrchSpecNewBtn.Name = "SrchSpecNewBtn";
            this.SrchSpecNewBtn.Size = new System.Drawing.Size(59, 23);
            this.SrchSpecNewBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.SrchSpecNewBtn.TabIndex = 4;
            this.SrchSpecNewBtn.TabStop = false;
            this.SrchSpecNewBtn.Text = "Обзор";
            this.SrchSpecNewBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.SrchSpecNewBtn.Click += new System.EventHandler(this.SrchSpecNewBtn_Click);
            // 
            // compareBtn
            // 
            this.compareBtn.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.compareBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.compareBtn.Location = new System.Drawing.Point(5, 358);
            this.compareBtn.Name = "compareBtn";
            this.compareBtn.Size = new System.Drawing.Size(388, 55);
            this.compareBtn.TabIndex = 6;
            this.compareBtn.Text = "Сравнить";
            this.compareBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.compareBtn.Click += new System.EventHandler(this.CompareBtn_Click);
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.Location = new System.Drawing.Point(55, 435);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(341, 57);
            this.metroLabel1.Style = MetroFramework.MetroColorStyle.Yellow;
            this.metroLabel1.TabIndex = 7;
            this.metroLabel1.Text = "Несоответствие данных\r\n (за истину приняты данные из новой спецификации,\r\nстарые " +
    "данные см. в примечании к ячейке)";
            this.metroLabel1.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroLabel1.UseStyleColors = true;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Red;
            this.panel1.Location = new System.Drawing.Point(6, 435);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(30, 19);
            this.panel1.TabIndex = 8;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Blue;
            this.panel2.Location = new System.Drawing.Point(6, 496);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(30, 19);
            this.panel2.TabIndex = 10;
            // 
            // metroLabel2
            // 
            this.metroLabel2.AutoSize = true;
            this.metroLabel2.Location = new System.Drawing.Point(55, 496);
            this.metroLabel2.Name = "metroLabel2";
            this.metroLabel2.Size = new System.Drawing.Size(99, 19);
            this.metroLabel2.Style = MetroFramework.MetroColorStyle.Yellow;
            this.metroLabel2.TabIndex = 9;
            this.metroLabel2.Text = "Повтор строки";
            this.metroLabel2.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroLabel2.UseStyleColors = true;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Yellow;
            this.panel3.Location = new System.Drawing.Point(6, 530);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(30, 19);
            this.panel3.TabIndex = 12;
            // 
            // metroLabel3
            // 
            this.metroLabel3.AutoSize = true;
            this.metroLabel3.Location = new System.Drawing.Point(55, 530);
            this.metroLabel3.Name = "metroLabel3";
            this.metroLabel3.Size = new System.Drawing.Size(271, 19);
            this.metroLabel3.Style = MetroFramework.MetroColorStyle.Yellow;
            this.metroLabel3.TabIndex = 11;
            this.metroLabel3.Text = "Строка отсутствует в новой спецификации";
            this.metroLabel3.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroLabel3.UseStyleColors = true;
            // 
            // metroLabel4
            // 
            this.metroLabel4.AutoSize = true;
            this.metroLabel4.Location = new System.Drawing.Point(55, 563);
            this.metroLabel4.Name = "metroLabel4";
            this.metroLabel4.Size = new System.Drawing.Size(274, 19);
            this.metroLabel4.Style = MetroFramework.MetroColorStyle.Yellow;
            this.metroLabel4.TabIndex = 11;
            this.metroLabel4.Text = "Строка отсутствует в старой спецификации";
            this.metroLabel4.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroLabel4.UseStyleColors = true;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.Green;
            this.panel4.Location = new System.Drawing.Point(6, 563);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(30, 19);
            this.panel4.TabIndex = 12;
            // 
            // metroLabel5
            // 
            this.metroLabel5.AutoSize = true;
            this.metroLabel5.Location = new System.Drawing.Point(104, 81);
            this.metroLabel5.Name = "metroLabel5";
            this.metroLabel5.Size = new System.Drawing.Size(189, 19);
            this.metroLabel5.Style = MetroFramework.MetroColorStyle.Yellow;
            this.metroLabel5.TabIndex = 13;
            this.metroLabel5.Text = "Номер колонки с артикулами";
            this.metroLabel5.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroLabel5.UseStyleColors = true;
            // 
            // artColTBox
            // 
            this.artColTBox.Location = new System.Drawing.Point(104, 103);
            this.artColTBox.Name = "artColTBox";
            this.artColTBox.Size = new System.Drawing.Size(189, 23);
            this.artColTBox.TabIndex = 14;
            this.artColTBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.artColTBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.artColTBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.columnsTBox_KeyPress);
            // 
            // headerChBox
            // 
            this.headerChBox.AutoSize = true;
            this.headerChBox.Location = new System.Drawing.Point(5, 63);
            this.headerChBox.Name = "headerChBox";
            this.headerChBox.Size = new System.Drawing.Size(157, 15);
            this.headerChBox.Style = MetroFramework.MetroColorStyle.Yellow;
            this.headerChBox.TabIndex = 15;
            this.headerChBox.Text = "Шапка в спецификациях";
            this.headerChBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.headerChBox.UseStyleColors = true;
            this.headerChBox.UseVisualStyleBackColor = true;
            // 
            // nameColTBox
            // 
            this.nameColTBox.Location = new System.Drawing.Point(105, 151);
            this.nameColTBox.Name = "nameColTBox";
            this.nameColTBox.Size = new System.Drawing.Size(189, 23);
            this.nameColTBox.TabIndex = 17;
            this.nameColTBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nameColTBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            // 
            // metroLabel6
            // 
            this.metroLabel6.AutoSize = true;
            this.metroLabel6.Location = new System.Drawing.Point(105, 129);
            this.metroLabel6.Name = "metroLabel6";
            this.metroLabel6.Size = new System.Drawing.Size(228, 19);
            this.metroLabel6.Style = MetroFramework.MetroColorStyle.Yellow;
            this.metroLabel6.TabIndex = 16;
            this.metroLabel6.Text = "Номер колонки с наименованиями";
            this.metroLabel6.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroLabel6.UseStyleColors = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 592);
            this.Controls.Add(this.nameColTBox);
            this.Controls.Add(this.metroLabel6);
            this.Controls.Add(this.headerChBox);
            this.Controls.Add(this.artColTBox);
            this.Controls.Add(this.metroLabel5);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.metroLabel4);
            this.Controls.Add(this.metroLabel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.metroLabel2);
            this.Controls.Add(this.metroLabel1);
            this.Controls.Add(this.compareBtn);
            this.Controls.Add(this.newVerTxt);
            this.Controls.Add(this.SrchSpecNewBtn);
            this.Controls.Add(this.oldVerTxt);
            this.Controls.Add(this.SrchSpecOldBtn);
            this.Controls.Add(this.oldVerLabel);
            this.Name = "MainForm";
            this.Resizable = false;
            this.Style = MetroFramework.MetroColorStyle.Yellow;
            this.Text = "SpecificationCompare";
            this.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MetroFramework.Controls.MetroLabel oldVerLabel;
        private MetroFramework.Controls.MetroButton SrchSpecOldBtn;
        private MetroFramework.Controls.MetroTextBox oldVerTxt;
        private MetroFramework.Controls.MetroTextBox newVerTxt;
        private MetroFramework.Controls.MetroButton SrchSpecNewBtn;
        private MetroFramework.Controls.MetroButton compareBtn;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private MetroFramework.Controls.MetroLabel metroLabel2;
        private System.Windows.Forms.Panel panel3;
        private MetroFramework.Controls.MetroLabel metroLabel3;
        private MetroFramework.Controls.MetroLabel metroLabel4;
        private System.Windows.Forms.Panel panel4;
        private MetroFramework.Controls.MetroLabel metroLabel5;
        private MetroFramework.Controls.MetroTextBox artColTBox;
        private MetroFramework.Controls.MetroCheckBox headerChBox;
        private MetroFramework.Controls.MetroTextBox nameColTBox;
        private MetroFramework.Controls.MetroLabel metroLabel6;
    }
}

