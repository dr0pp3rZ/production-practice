namespace obzor
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            dataGridView1 = new DataGridView();
            menuStrip1 = new MenuStrip();
            файлToolStripMenuItem = new ToolStripMenuItem();
            открытьToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripMenuItem();
            toolStripMenuItem2 = new ToolStripMenuItem();
            toolStripMenuItem3 = new ToolStripMenuItem();
            toolStrip1 = new ToolStrip();
            toolStripLabel1 = new ToolStripLabel();
            toolStripComboBox1 = new ToolStripComboBox();
            openFileDialog1 = new OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            menuStrip1.SuspendLayout();
            toolStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.Location = new Point(0, 24);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(879, 524);
            dataGridView1.TabIndex = 4;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { файлToolStripMenuItem, toolStripMenuItem1, toolStripMenuItem2, toolStripMenuItem3 });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(879, 24);
            menuStrip1.TabIndex = 5;
            menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            файлToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { открытьToolStripMenuItem });
            файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            файлToolStripMenuItem.Size = new Size(48, 20);
            файлToolStripMenuItem.Text = "Файл";
            // 
            // открытьToolStripMenuItem
            // 
            открытьToolStripMenuItem.Name = "открытьToolStripMenuItem";
            открытьToolStripMenuItem.Size = new Size(128, 22);
            открытьToolStripMenuItem.Text = "Загрузить";
            открытьToolStripMenuItem.Click += открытьToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(71, 20);
            toolStripMenuItem1.Text = "Добавить";
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(73, 20);
            toolStripMenuItem2.Text = "Изменить";
            // 
            // toolStripMenuItem3
            // 
            toolStripMenuItem3.Name = "toolStripMenuItem3";
            toolStripMenuItem3.Size = new Size(63, 20);
            toolStripMenuItem3.Text = "Удалить";
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripLabel1, toolStripComboBox1 });
            toolStrip1.Location = new Point(0, 24);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(879, 25);
            toolStrip1.TabIndex = 6;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            toolStripLabel1.Name = "toolStripLabel1";
            toolStripLabel1.Size = new Size(33, 22);
            toolStripLabel1.Text = "Лист";
            // 
            // toolStripComboBox1
            // 
            toolStripComboBox1.Name = "toolStripComboBox1";
            toolStripComboBox1.Size = new Size(121, 25);
            toolStripComboBox1.SelectedIndexChanged += toolStripComboBox1_SelectedIndexChanged;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            openFileDialog1.Filter = "Excel|*.xlsx";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(879, 548);
            Controls.Add(toolStrip1);
            Controls.Add(dataGridView1);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Загрузка файла";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private DataGridView dataGridView1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem файлToolStripMenuItem;
        private ToolStripMenuItem открытьToolStripMenuItem;
        private ToolStrip toolStrip1;
        private ToolStripLabel toolStripLabel1;
        private ToolStripComboBox toolStripComboBox1;
        private OpenFileDialog openFileDialog1;
        private ToolStripMenuItem toolStripMenuItem1;
        private ToolStripMenuItem toolStripMenuItem2;
        private ToolStripMenuItem toolStripMenuItem3;
    }
}