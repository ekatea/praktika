namespace WindowsFormsApp1
{
    partial class frm
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
            this.btn1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn1
            // 
            this.btn1.Location = new System.Drawing.Point(45, 34);
            this.btn1.Name = "btn1";
            this.btn1.Size = new System.Drawing.Size(207, 91);
            this.btn1.TabIndex = 0;
            this.btn1.Tag = "btn1";
            this.btn1.Text = "Создать Word";
            this.btn1.UseVisualStyleBackColor = true;
            this.btn1.Click += new System.EventHandler(this.btn1_Click);
            // 
            // frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn1);
            this.Name = "frm";
            this.Text = "Задание 1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn1;
    }
}

