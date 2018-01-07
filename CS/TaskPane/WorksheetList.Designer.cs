namespace Favorites.TaskPane
{
    partial class WorksheetList
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WorksheetList));
            this.imgXlSheetVisibility = new System.Windows.Forms.ImageList(this.components);
            this.mnuSetVisiblity = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tmiVisible = new System.Windows.Forms.ToolStripMenuItem();
            this.tmiHidden = new System.Windows.Forms.ToolStripMenuItem();
            this.tmiVeryHidden = new System.Windows.Forms.ToolStripMenuItem();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lstWorksheets = new System.Windows.Forms.ListView();
            this.tspWorksheetMenu = new System.Windows.Forms.ToolStrip();
            this.tsbRefresh = new System.Windows.Forms.ToolStripButton();
            this.mnuSetVisiblity.SuspendLayout();
            this.tspWorksheetMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // imgXlSheetVisibility
            // 
            this.imgXlSheetVisibility.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgXlSheetVisibility.ImageStream")));
            this.imgXlSheetVisibility.TransparentColor = System.Drawing.Color.Transparent;
            this.imgXlSheetVisibility.Images.SetKeyName(0, "bullet_green.png");
            this.imgXlSheetVisibility.Images.SetKeyName(1, "bullet_red.png");
            this.imgXlSheetVisibility.Images.SetKeyName(2, "bullet_key.png");
            // 
            // mnuSetVisiblity
            // 
            this.mnuSetVisiblity.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tmiVisible,
            this.tmiHidden,
            this.tmiVeryHidden});
            this.mnuSetVisiblity.Name = "contextMenuStrip1";
            this.mnuSetVisiblity.Size = new System.Drawing.Size(140, 70);
            // 
            // tmiVisible
            // 
            this.tmiVisible.Image = global::Favorites.Properties.Resources.bullet_green;
            this.tmiVisible.Name = "tmiVisible";
            this.tmiVisible.Size = new System.Drawing.Size(139, 22);
            this.tmiVisible.Text = "Visible";
            this.tmiVisible.Click += new System.EventHandler(this.tmiVisiblity_Click);
            // 
            // tmiHidden
            // 
            this.tmiHidden.Image = global::Favorites.Properties.Resources.bullet_red;
            this.tmiHidden.Name = "tmiHidden";
            this.tmiHidden.Size = new System.Drawing.Size(139, 22);
            this.tmiHidden.Text = "Hidden";
            this.tmiHidden.Click += new System.EventHandler(this.tmiVisiblity_Click);
            // 
            // tmiVeryHidden
            // 
            this.tmiVeryHidden.Image = global::Favorites.Properties.Resources.bullet_key;
            this.tmiVeryHidden.Name = "tmiVeryHidden";
            this.tmiVeryHidden.Size = new System.Drawing.Size(139, 22);
            this.tmiVeryHidden.Text = "Very Hidden";
            this.tmiVeryHidden.Click += new System.EventHandler(this.tmiVisiblity_Click);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Worksheet Name";
            this.columnHeader1.Width = 500;
            // 
            // lstWorksheets
            // 
            this.lstWorksheets.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstWorksheets.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.lstWorksheets.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstWorksheets.FullRowSelect = true;
            this.lstWorksheets.GridLines = true;
            this.lstWorksheets.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.lstWorksheets.LabelWrap = false;
            this.lstWorksheets.Location = new System.Drawing.Point(0, 28);
            this.lstWorksheets.MultiSelect = false;
            this.lstWorksheets.Name = "lstWorksheets";
            this.lstWorksheets.Size = new System.Drawing.Size(300, 722);
            this.lstWorksheets.TabIndex = 2;
            this.lstWorksheets.UseCompatibleStateImageBehavior = false;
            this.lstWorksheets.View = System.Windows.Forms.View.Details;
            this.lstWorksheets.SelectedIndexChanged += new System.EventHandler(this.lstWorksheets_SelectedIndexChanged);
            this.lstWorksheets.MouseClick += new System.Windows.Forms.MouseEventHandler(this.lstWorksheets_MouseClick);
            // 
            // tspWorksheetMenu
            // 
            this.tspWorksheetMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.tspWorksheetMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbRefresh});
            this.tspWorksheetMenu.Location = new System.Drawing.Point(0, 0);
            this.tspWorksheetMenu.Name = "tspWorksheetMenu";
            this.tspWorksheetMenu.Size = new System.Drawing.Size(300, 25);
            this.tspWorksheetMenu.TabIndex = 3;
            this.tspWorksheetMenu.Text = "toolStrip1";
            // 
            // tsbRefresh
            // 
            this.tsbRefresh.Image = global::Favorites.Properties.Resources.arrow_refresh;
            this.tsbRefresh.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbRefresh.Name = "tsbRefresh";
            this.tsbRefresh.Size = new System.Drawing.Size(69, 22);
            this.tsbRefresh.Text = " Refresh";
            this.tsbRefresh.Click += new System.EventHandler(this.tsbRefresh_Click);
            // 
            // WorksheetList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tspWorksheetMenu);
            this.Controls.Add(this.lstWorksheets);
            this.MinimumSize = new System.Drawing.Size(300, 750);
            this.Name = "WorksheetList";
            this.Size = new System.Drawing.Size(300, 750);
            this.Load += new System.EventHandler(this.WorksheetList_Load);
            this.mnuSetVisiblity.ResumeLayout(false);
            this.tspWorksheetMenu.ResumeLayout(false);
            this.tspWorksheetMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ImageList imgXlSheetVisibility;
        private System.Windows.Forms.ContextMenuStrip mnuSetVisiblity;
        private System.Windows.Forms.ToolStripMenuItem tmiVisible;
        private System.Windows.Forms.ToolStripMenuItem tmiHidden;
        private System.Windows.Forms.ToolStripMenuItem tmiVeryHidden;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ListView lstWorksheets;
        private System.Windows.Forms.ToolStrip tspWorksheetMenu;
        private System.Windows.Forms.ToolStripButton tsbRefresh;
    }
}
