namespace LimsDesign
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.bgw = new System.ComponentModel.BackgroundWorker();
            this.ts = new System.Windows.Forms.ToolStrip();
            this.ss = new System.Windows.Forms.StatusStrip();
            this.tspb = new System.Windows.Forms.ToolStripProgressBar();
            this.tssl = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsbtnExportTableDefine2Markdown = new System.Windows.Forms.ToolStripButton();
            this.tsbtnExportTableDefine2Word = new System.Windows.Forms.ToolStripButton();
            this.ts.SuspendLayout();
            this.ss.SuspendLayout();
            this.SuspendLayout();
            // 
            // bgw
            // 
            this.bgw.WorkerReportsProgress = true;
            this.bgw.WorkerSupportsCancellation = true;
            this.bgw.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgw_DoWork);
            this.bgw.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgw_ProgressChanged);
            this.bgw.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgw_RunWorkerCompleted);
            // 
            // ts
            // 
            this.ts.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.ts.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbtnExportTableDefine2Markdown,
            this.tsbtnExportTableDefine2Word});
            this.ts.Location = new System.Drawing.Point(0, 0);
            this.ts.Name = "ts";
            this.ts.Size = new System.Drawing.Size(520, 27);
            this.ts.TabIndex = 2;
            this.ts.Text = "toolStrip1";
            // 
            // ss
            // 
            this.ss.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.ss.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tssl,
            this.tspb});
            this.ss.Location = new System.Drawing.Point(0, 56);
            this.ss.Name = "ss";
            this.ss.Size = new System.Drawing.Size(520, 25);
            this.ss.TabIndex = 3;
            this.ss.Text = "statusStrip1";
            // 
            // tspb
            // 
            this.tspb.Name = "tspb";
            this.tspb.Size = new System.Drawing.Size(500, 19);
            // 
            // tssl
            // 
            this.tssl.Name = "tssl";
            this.tssl.Size = new System.Drawing.Size(0, 20);
            // 
            // tsbtnExportTableDefine2Markdown
            // 
            this.tsbtnExportTableDefine2Markdown.Image = ((System.Drawing.Image)(resources.GetObject("tsbtnExportTableDefine2Markdown.Image")));
            this.tsbtnExportTableDefine2Markdown.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbtnExportTableDefine2Markdown.Name = "tsbtnExportTableDefine2Markdown";
            this.tsbtnExportTableDefine2Markdown.Size = new System.Drawing.Size(156, 24);
            this.tsbtnExportTableDefine2Markdown.Text = "导出到Markdown";
            this.tsbtnExportTableDefine2Markdown.Click += new System.EventHandler(this.tsbtnExportTableDefine2Markdown_Click);
            // 
            // tsbtnExportTableDefine2Word
            // 
            this.tsbtnExportTableDefine2Word.Image = ((System.Drawing.Image)(resources.GetObject("tsbtnExportTableDefine2Word.Image")));
            this.tsbtnExportTableDefine2Word.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbtnExportTableDefine2Word.Name = "tsbtnExportTableDefine2Word";
            this.tsbtnExportTableDefine2Word.Size = new System.Drawing.Size(119, 24);
            this.tsbtnExportTableDefine2Word.Text = "导出到Word";
            this.tsbtnExportTableDefine2Word.Click += new System.EventHandler(this.tsbtnExportTableDefine2Word_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(520, 81);
            this.Controls.Add(this.ss);
            this.Controls.Add(this.ts);
            this.Name = "Form1";
            this.Text = "starlims表结构导出";
            this.ts.ResumeLayout(false);
            this.ts.PerformLayout();
            this.ss.ResumeLayout(false);
            this.ss.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.ComponentModel.BackgroundWorker bgw;
        private System.Windows.Forms.ToolStrip ts;
        private System.Windows.Forms.StatusStrip ss;
        private System.Windows.Forms.ToolStripProgressBar tspb;
        private System.Windows.Forms.ToolStripStatusLabel tssl;
        private System.Windows.Forms.ToolStripButton tsbtnExportTableDefine2Markdown;
        private System.Windows.Forms.ToolStripButton tsbtnExportTableDefine2Word;
    }
}

