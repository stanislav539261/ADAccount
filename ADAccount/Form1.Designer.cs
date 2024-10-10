namespace ADAccount
{
    partial class ADAccount
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADAccount));
            this.label2 = new System.Windows.Forms.Label();
            this.chboxEncryption = new System.Windows.Forms.CheckBox();
            this.chboxForbidChange = new System.Windows.Forms.CheckBox();
            this.chboxNeverExpires = new System.Windows.Forms.CheckBox();
            this.table = new System.Windows.Forms.DataGridView();
            this.name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.group = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dob = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnLoad = new System.Windows.Forms.Button();
            this.txboxPasswd = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txboxDomain = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.chboxGenerate = new System.Windows.Forms.CheckBox();
            this.radioAdd = new System.Windows.Forms.RadioButton();
            this.btnApply = new System.Windows.Forms.Button();
            this.radioDelete = new System.Windows.Forms.RadioButton();
            this.chboxOUGroup = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.table)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(737, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(435, 44);
            this.label2.TabIndex = 0;
            this.label2.Text = "Параметры учетной записи:";
            // 
            // chboxEncryption
            // 
            this.chboxEncryption.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chboxEncryption.Location = new System.Drawing.Point(742, 195);
            this.chboxEncryption.Name = "chboxEncryption";
            this.chboxEncryption.Size = new System.Drawing.Size(429, 64);
            this.chboxEncryption.TabIndex = 1;
            this.chboxEncryption.Text = "Хранить пароль, используя обратимое шифрование";
            this.chboxEncryption.UseVisualStyleBackColor = true;
            // 
            // chboxForbidChange
            // 
            this.chboxForbidChange.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chboxForbidChange.Location = new System.Drawing.Point(742, 119);
            this.chboxForbidChange.Name = "chboxForbidChange";
            this.chboxForbidChange.Size = new System.Drawing.Size(429, 32);
            this.chboxForbidChange.TabIndex = 2;
            this.chboxForbidChange.Text = "Запретить смену пароля пользователем";
            this.chboxForbidChange.UseVisualStyleBackColor = true;
            // 
            // chboxNeverExpires
            // 
            this.chboxNeverExpires.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chboxNeverExpires.Location = new System.Drawing.Point(742, 157);
            this.chboxNeverExpires.Name = "chboxNeverExpires";
            this.chboxNeverExpires.Size = new System.Drawing.Size(429, 32);
            this.chboxNeverExpires.TabIndex = 3;
            this.chboxNeverExpires.Text = "Срок действия пароля не ограничен";
            this.chboxNeverExpires.UseVisualStyleBackColor = true;
            // 
            // table
            // 
            this.table.AllowUserToAddRows = false;
            this.table.AllowUserToDeleteRows = false;
            this.table.AllowUserToResizeColumns = false;
            this.table.AllowUserToResizeRows = false;
            this.table.BackgroundColor = System.Drawing.SystemColors.Control;
            this.table.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.table.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.table.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.table.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.name,
            this.group,
            this.dob,
            this.ad});
            this.table.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.table.EnableHeadersVisualStyles = false;
            this.table.GridColor = System.Drawing.SystemColors.ControlDarkDark;
            this.table.Location = new System.Drawing.Point(12, 72);
            this.table.Name = "table";
            this.table.ReadOnly = true;
            this.table.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.table.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.table.RowHeadersVisible = false;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.table.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.table.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.table.Size = new System.Drawing.Size(719, 577);
            this.table.TabIndex = 4;
            // 
            // name
            // 
            this.name.HeaderText = "ФИО";
            this.name.Name = "name";
            this.name.ReadOnly = true;
            this.name.Width = 350;
            // 
            // group
            // 
            this.group.HeaderText = "Группа";
            this.group.Name = "group";
            this.group.ReadOnly = true;
            this.group.Width = 80;
            // 
            // dob
            // 
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dob.DefaultCellStyle = dataGridViewCellStyle2;
            this.dob.HeaderText = "Дата рождения";
            this.dob.Name = "dob";
            this.dob.ReadOnly = true;
            this.dob.Width = 155;
            // 
            // ad
            // 
            this.ad.HeaderText = "SAM";
            this.ad.Name = "ad";
            this.ad.ReadOnly = true;
            this.ad.Width = 133;
            // 
            // btnLoad
            // 
            this.btnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnLoad.Location = new System.Drawing.Point(980, 8);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(192, 48);
            this.btnLoad.TabIndex = 5;
            this.btnLoad.Text = "Загрузить Excel";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.OpenExcelFile);
            // 
            // txboxPasswd
            // 
            this.txboxPasswd.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txboxPasswd.Location = new System.Drawing.Point(820, 396);
            this.txboxPasswd.Name = "txboxPasswd";
            this.txboxPasswd.Size = new System.Drawing.Size(242, 29);
            this.txboxPasswd.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(737, 399);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 24);
            this.label1.TabIndex = 8;
            this.label1.Text = "Пароль";
            // 
            // txboxDomain
            // 
            this.txboxDomain.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txboxDomain.Location = new System.Drawing.Point(88, 20);
            this.txboxDomain.Name = "txboxDomain";
            this.txboxDomain.Size = new System.Drawing.Size(355, 29);
            this.txboxDomain.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(12, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 24);
            this.label3.TabIndex = 10;
            this.label3.Text = "Домен";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ADAccount.Properties.Resources.settings;
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(449, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(44, 44);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 14;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.Settings);
            this.pictureBox1.MouseLeave += new System.EventHandler(this.pictureBox1_MouseLeave);
            this.pictureBox1.MouseHover += new System.EventHandler(this.pictureBox1_MouseHover);
            // 
            // chboxGenerate
            // 
            this.chboxGenerate.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chboxGenerate.Location = new System.Drawing.Point(741, 265);
            this.chboxGenerate.Name = "chboxGenerate";
            this.chboxGenerate.Size = new System.Drawing.Size(429, 64);
            this.chboxGenerate.TabIndex = 15;
            this.chboxGenerate.Text = "Сгенерировать пароль на основе даты рождения пользователя (фам_ммгг)";
            this.chboxGenerate.UseVisualStyleBackColor = true;
            this.chboxGenerate.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // radioAdd
            // 
            this.radioAdd.AutoSize = true;
            this.radioAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.radioAdd.Location = new System.Drawing.Point(742, 612);
            this.radioAdd.Name = "radioAdd";
            this.radioAdd.Size = new System.Drawing.Size(117, 28);
            this.radioAdd.TabIndex = 16;
            this.radioAdd.TabStop = true;
            this.radioAdd.Text = "Добавить";
            this.radioAdd.UseVisualStyleBackColor = true;
            this.radioAdd.CheckedChanged += new System.EventHandler(this.RadioButtonsChanged);
            // 
            // btnApply
            // 
            this.btnApply.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnApply.Location = new System.Drawing.Point(992, 601);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(180, 48);
            this.btnApply.TabIndex = 17;
            this.btnApply.Text = "Применить";
            this.btnApply.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.ApplyChanges);
            // 
            // radioDelete
            // 
            this.radioDelete.AutoSize = true;
            this.radioDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.radioDelete.Location = new System.Drawing.Point(865, 612);
            this.radioDelete.Name = "radioDelete";
            this.radioDelete.Size = new System.Drawing.Size(104, 28);
            this.radioDelete.TabIndex = 18;
            this.radioDelete.TabStop = true;
            this.radioDelete.Text = "Удалить";
            this.radioDelete.UseVisualStyleBackColor = true;
            this.radioDelete.CheckedChanged += new System.EventHandler(this.RadioButtonsChanged);
            // 
            // chboxOUGroup
            // 
            this.chboxOUGroup.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chboxOUGroup.Location = new System.Drawing.Point(741, 335);
            this.chboxOUGroup.Name = "chboxOUGroup";
            this.chboxOUGroup.Size = new System.Drawing.Size(429, 32);
            this.chboxOUGroup.TabIndex = 19;
            this.chboxOUGroup.Text = "Подразделения OU вместо групп";
            this.chboxOUGroup.UseVisualStyleBackColor = true;
            // 
            // ADAccount
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1184, 661);
            this.Controls.Add(this.chboxOUGroup);
            this.Controls.Add(this.radioDelete);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.radioAdd);
            this.Controls.Add(this.chboxGenerate);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txboxDomain);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txboxPasswd);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.table);
            this.Controls.Add(this.chboxNeverExpires);
            this.Controls.Add(this.chboxForbidChange);
            this.Controls.Add(this.chboxEncryption);
            this.Controls.Add(this.label2);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ADAccount";
            this.Text = "ADAccount";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.DragDropWindow);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.DragEnterWindow);
            this.Resize += new System.EventHandler(this.ResizeWindow);
            ((System.ComponentModel.ISupportInitialize)(this.table)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chboxEncryption;
        private System.Windows.Forms.CheckBox chboxForbidChange;
        private System.Windows.Forms.CheckBox chboxNeverExpires;
        private System.Windows.Forms.DataGridView table;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.TextBox txboxPasswd;
        private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txboxDomain;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.CheckBox chboxGenerate;
		private System.Windows.Forms.RadioButton radioAdd;
		private System.Windows.Forms.Button btnApply;
		private System.Windows.Forms.RadioButton radioDelete;
		private System.Windows.Forms.DataGridViewTextBoxColumn name;
		private System.Windows.Forms.DataGridViewTextBoxColumn group;
		private System.Windows.Forms.DataGridViewTextBoxColumn dob;
		private System.Windows.Forms.DataGridViewTextBoxColumn ad;
		private System.Windows.Forms.CheckBox chboxOUGroup;
	}
}

