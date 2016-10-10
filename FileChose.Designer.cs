namespace MyApp
{
    partial class FileChose
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FileChose));
            this.line1 = new DevComponents.DotNetBar.Controls.Line();
            this.label3 = new System.Windows.Forms.Label();
            this.maskedTextBoxAdv1 = new DevComponents.DotNetBar.Controls.MaskedTextBoxAdv();
            this.buttonX2 = new DevComponents.DotNetBar.ButtonX();
            this.buttonX4 = new DevComponents.DotNetBar.ButtonX();
            this.buttonX5 = new DevComponents.DotNetBar.ButtonX();
            this.textBoxX3 = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.radio1 = new System.Windows.Forms.RadioButton();
            this.radio2 = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // line1
            // 
            this.line1.Location = new System.Drawing.Point(-2, 167);
            this.line1.Name = "line1";
            this.line1.Size = new System.Drawing.Size(602, 23);
            this.line1.TabIndex = 3;
            this.line1.Text = "line1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(137, 193);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 16);
            this.label3.TabIndex = 4;
            this.label3.Text = "记不清了？";
            // 
            // maskedTextBoxAdv1
            // 
            // 
            // 
            // 
            this.maskedTextBoxAdv1.BackgroundStyle.Class = "TextBoxBorder";
            this.maskedTextBoxAdv1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.maskedTextBoxAdv1.ButtonClear.Visible = true;
            this.maskedTextBoxAdv1.Location = new System.Drawing.Point(201, 224);
            this.maskedTextBoxAdv1.Name = "maskedTextBoxAdv1";
            this.maskedTextBoxAdv1.Size = new System.Drawing.Size(151, 20);
            this.maskedTextBoxAdv1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.maskedTextBoxAdv1.TabIndex = 5;
            this.maskedTextBoxAdv1.Text = "";
            // 
            // buttonX2
            // 
            this.buttonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("buttonX2.BackgroundImage")));
            this.buttonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.buttonX2.HotTrackingStyle = DevComponents.DotNetBar.eHotTrackingStyle.None;
            this.buttonX2.Location = new System.Drawing.Point(353, 135);
            this.buttonX2.Name = "buttonX2";
            this.buttonX2.Size = new System.Drawing.Size(65, 31);
            this.buttonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2013;
            this.buttonX2.TabIndex = 6;
            this.buttonX2.Text = "书式管理";
            this.buttonX2.Click += new System.EventHandler(this.buttonX2_Click);
            // 
            // buttonX4
            // 
            this.buttonX4.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX4.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("buttonX4.BackgroundImage")));
            this.buttonX4.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.buttonX4.Enabled = false;
            this.buttonX4.Font = new System.Drawing.Font("楷体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonX4.Location = new System.Drawing.Point(140, 12);
            this.buttonX4.Name = "buttonX4";
            this.buttonX4.Size = new System.Drawing.Size(326, 30);
            this.buttonX4.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX4.TabIndex = 7;
            this.buttonX4.Text = "商标书式管理系统";
            // 
            // buttonX5
            // 
            this.buttonX5.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX5.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("buttonX5.BackgroundImage")));
            this.buttonX5.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.buttonX5.Location = new System.Drawing.Point(353, 262);
            this.buttonX5.Name = "buttonX5";
            this.buttonX5.Size = new System.Drawing.Size(65, 31);
            this.buttonX5.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX5.TabIndex = 8;
            this.buttonX5.Text = "查询";
            this.buttonX5.TextColor = System.Drawing.Color.White;
            this.buttonX5.Click += new System.EventHandler(this.buttonX5_Click);
            // 
            // textBoxX3
            // 
            this.textBoxX3.BackColor = System.Drawing.Color.White;
            // 
            // 
            // 
            this.textBoxX3.Border.Class = "TextBoxBorder";
            this.textBoxX3.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.textBoxX3.DisabledBackColor = System.Drawing.Color.White;
            this.textBoxX3.ForeColor = System.Drawing.Color.Black;
            this.textBoxX3.Location = new System.Drawing.Point(328, 108);
            this.textBoxX3.Name = "textBoxX3";
            this.textBoxX3.PreventEnterBeep = true;
            this.textBoxX3.Size = new System.Drawing.Size(100, 21);
            this.textBoxX3.TabIndex = 9;
            this.textBoxX3.TextChanged += new System.EventHandler(this.textBoxX3_TextChanged);
            // 
            // radio1
            // 
            this.radio1.AutoSize = true;
            this.radio1.Location = new System.Drawing.Point(178, 73);
            this.radio1.Name = "radio1";
            this.radio1.Size = new System.Drawing.Size(59, 16);
            this.radio1.TabIndex = 10;
            this.radio1.TabStop = true;
            this.radio1.Text = "新客户";
            this.radio1.UseVisualStyleBackColor = true;
            this.radio1.CheckedChanged += new System.EventHandler(this.radio1_CheckedChanged);
            // 
            // radio2
            // 
            this.radio2.AutoSize = true;
            this.radio2.Location = new System.Drawing.Point(284, 73);
            this.radio2.Name = "radio2";
            this.radio2.Size = new System.Drawing.Size(59, 16);
            this.radio2.TabIndex = 11;
            this.radio2.TabStop = true;
            this.radio2.Text = "老客户";
            this.radio2.UseVisualStyleBackColor = true;
            this.radio2.CheckedChanged += new System.EventHandler(this.radio2_CheckedChanged);
            // 
            // FileChose
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(61)))), ((int)(((byte)(177)))), ((int)(((byte)(158)))));
            this.ClientSize = new System.Drawing.Size(595, 373);
            this.Controls.Add(this.radio2);
            this.Controls.Add(this.radio1);
            this.Controls.Add(this.textBoxX3);
            this.Controls.Add(this.buttonX5);
            this.Controls.Add(this.buttonX4);
            this.Controls.Add(this.buttonX2);
            this.Controls.Add(this.maskedTextBoxAdv1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.line1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(49)))), ((int)(((byte)(59)))));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FileChose";
            this.Text = "登入客户";
            this.Load += new System.EventHandler(this.FileChose_Load_1);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX buttonX3;
        private DevComponents.DotNetBar.TabControl tabControl1;
        private DevComponents.DotNetBar.TabControlPanel tabControlPanel1;
        private DevComponents.DotNetBar.TabItem tabItem6;
        private DevComponents.DotNetBar.TabControlPanel tabControlPanel2;
        private DevComponents.DotNetBar.TabItem tabItem1;
        private DevComponents.DotNetBar.TabControlPanel tabControlPanel3;
        private DevComponents.DotNetBar.TabItem tabItem2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.TextBoxX textBoxX1;
        private DevComponents.DotNetBar.ButtonX buttonX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.TextBoxX textBoxX2;
        private DevComponents.DotNetBar.Controls.DataGridViewX dataGridViewX1;
        private DevComponents.DotNetBar.Controls.Line line1;
        private System.Windows.Forms.Label label3;
        private DevComponents.DotNetBar.Controls.MaskedTextBoxAdv maskedTextBoxAdv1;
        private DevComponents.DotNetBar.ButtonX buttonX2;
        private DevComponents.DotNetBar.ButtonX buttonX4;
        private DevComponents.DotNetBar.ButtonX buttonX5;
        private DevComponents.DotNetBar.Controls.TextBoxX textBoxX3;
        private System.Windows.Forms.RadioButton radio1;
        private System.Windows.Forms.RadioButton radio2;

    }
}