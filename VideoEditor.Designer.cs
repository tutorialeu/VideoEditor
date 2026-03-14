
namespace VideoEditor
{
    partial class VideoEditor
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VideoEditor));
            this.pnVideoEditing = new System.Windows.Forms.Panel();
            this.pbCursor = new System.Windows.Forms.PictureBox();
            this.hScrollBarZoom = new System.Windows.Forms.HScrollBar();
            this.cursorTimer = new System.Windows.Forms.Timer(this.components);
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.pnElements = new System.Windows.Forms.Panel();
            this.pnTitle = new System.Windows.Forms.Panel();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.btnMinimize = new System.Windows.Forms.Button();
            this.btnMaximize = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ediToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.undoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.redoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.cutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.renderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.createVideoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.infoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitVertical = new System.Windows.Forms.SplitContainer();
            this.axWindowsMediaPlayer1 = new AxWMPLib.AxWindowsMediaPlayer();
            this.splitHorizontal = new System.Windows.Forms.SplitContainer();
            this.pnLeftMenu = new System.Windows.Forms.Panel();
            this.fileSystemWatcher = new System.IO.FileSystemWatcher();
            this.pnVideoEditing.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbCursor)).BeginInit();
            this.pnElements.SuspendLayout();
            this.pnTitle.SuspendLayout();
            this.menuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitVertical)).BeginInit();
            this.splitVertical.Panel1.SuspendLayout();
            this.splitVertical.Panel2.SuspendLayout();
            this.splitVertical.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axWindowsMediaPlayer1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitHorizontal)).BeginInit();
            this.splitHorizontal.Panel1.SuspendLayout();
            this.splitHorizontal.Panel2.SuspendLayout();
            this.splitHorizontal.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher)).BeginInit();
            this.SuspendLayout();
            // 
            // pnVideoEditing
            // 
            this.pnVideoEditing.AllowDrop = true;
            this.pnVideoEditing.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.pnVideoEditing.AutoScroll = true;
            this.pnVideoEditing.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.pnVideoEditing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnVideoEditing.Controls.Add(this.pbCursor);
            this.pnVideoEditing.Location = new System.Drawing.Point(51, 0);
            this.pnVideoEditing.Name = "pnVideoEditing";
            this.pnVideoEditing.Size = new System.Drawing.Size(1195, 264);
            this.pnVideoEditing.TabIndex = 0;
            this.pnVideoEditing.DragDrop += new System.Windows.Forms.DragEventHandler(this.pnVideoEditing_DragDrop);
            this.pnVideoEditing.DragEnter += new System.Windows.Forms.DragEventHandler(this.pnVideoEditing_DragEnter);
            this.pnVideoEditing.DragLeave += new System.EventHandler(this.pnVideoEditing_DragLeave);
            this.pnVideoEditing.Paint += new System.Windows.Forms.PaintEventHandler(this.pnVideoEditing_Paint);
            this.pnVideoEditing.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pnVideoEditing_MouseDown);
            this.pnVideoEditing.MouseLeave += new System.EventHandler(this.pnVideoEditing_MouseLeave);
            this.pnVideoEditing.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pnVideoEditing_MouseMove);
            this.pnVideoEditing.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pnVideoEditing_MouseUp);
            this.pnVideoEditing.Resize += new System.EventHandler(this.pnVideoEditing_Resize);
            // 
            // pbCursor
            // 
            this.pbCursor.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.pbCursor.Cursor = System.Windows.Forms.Cursors.VSplit;
            this.pbCursor.Location = new System.Drawing.Point(51, 0);
            this.pbCursor.Name = "pbCursor";
            this.pbCursor.Size = new System.Drawing.Size(2, 250);
            this.pbCursor.TabIndex = 0;
            this.pbCursor.TabStop = false;
            this.pbCursor.Click += new System.EventHandler(this.pbCursor_Click);
            this.pbCursor.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pbCursor_MouseMove);
            this.pbCursor.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pbCursor_MouseUp);
            // 
            // hScrollBarZoom
            // 
            this.hScrollBarZoom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.hScrollBarZoom.LargeChange = 1;
            this.hScrollBarZoom.Location = new System.Drawing.Point(-1, 272);
            this.hScrollBarZoom.Maximum = 5;
            this.hScrollBarZoom.Name = "hScrollBarZoom";
            this.hScrollBarZoom.Size = new System.Drawing.Size(212, 12);
            this.hScrollBarZoom.TabIndex = 50;
            this.hScrollBarZoom.Tag = "0";
            this.hScrollBarZoom.ValueChanged += new System.EventHandler(this.hScrollBarZoom_ValueChanged);
            // 
            // cursorTimer
            // 
            this.cursorTimer.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // pnElements
            // 
            this.pnElements.BackColor = System.Drawing.Color.White;
            this.pnElements.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnElements.Controls.Add(this.hScrollBarZoom);
            this.pnElements.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnElements.Location = new System.Drawing.Point(0, 0);
            this.pnElements.Name = "pnElements";
            this.pnElements.Size = new System.Drawing.Size(618, 286);
            this.pnElements.TabIndex = 3;
            // 
            // pnTitle
            // 
            this.pnTitle.BackColor = System.Drawing.Color.Transparent;
            this.pnTitle.Controls.Add(this.lblProjectName);
            this.pnTitle.Controls.Add(this.btnMinimize);
            this.pnTitle.Controls.Add(this.btnMaximize);
            this.pnTitle.Controls.Add(this.btnExit);
            this.pnTitle.Controls.Add(this.menuStrip);
            this.pnTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnTitle.Location = new System.Drawing.Point(0, 0);
            this.pnTitle.Margin = new System.Windows.Forms.Padding(0);
            this.pnTitle.Name = "pnTitle";
            this.pnTitle.Size = new System.Drawing.Size(1247, 28);
            this.pnTitle.TabIndex = 4;
            this.pnTitle.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.pnTitle_MouseDoubleClick);
            this.pnTitle.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pnTitle_MouseDown);
            // 
            // lblProjectName
            // 
            this.lblProjectName.AutoSize = true;
            this.lblProjectName.Location = new System.Drawing.Point(527, 9);
            this.lblProjectName.Name = "lblProjectName";
            this.lblProjectName.Size = new System.Drawing.Size(0, 13);
            this.lblProjectName.TabIndex = 3;
            this.lblProjectName.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lblProjectName_MouseDown);
            // 
            // btnMinimize
            // 
            this.btnMinimize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMinimize.BackColor = System.Drawing.Color.Transparent;
            this.btnMinimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMinimize.ForeColor = System.Drawing.Color.Black;
            this.btnMinimize.Location = new System.Drawing.Point(1170, 1);
            this.btnMinimize.Margin = new System.Windows.Forms.Padding(0);
            this.btnMinimize.Name = "btnMinimize";
            this.btnMinimize.Size = new System.Drawing.Size(24, 24);
            this.btnMinimize.TabIndex = 2;
            this.btnMinimize.TabStop = false;
            this.btnMinimize.Text = "🗕";
            this.btnMinimize.UseVisualStyleBackColor = false;
            this.btnMinimize.Click += new System.EventHandler(this.btnMinimize_Click);
            // 
            // btnMaximize
            // 
            this.btnMaximize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMaximize.BackColor = System.Drawing.Color.Transparent;
            this.btnMaximize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMaximize.Location = new System.Drawing.Point(1196, 1);
            this.btnMaximize.Name = "btnMaximize";
            this.btnMaximize.Size = new System.Drawing.Size(24, 24);
            this.btnMaximize.TabIndex = 1;
            this.btnMaximize.TabStop = false;
            this.btnMaximize.Text = "🗖";
            this.btnMaximize.UseVisualStyleBackColor = false;
            this.btnMaximize.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnExit
            // 
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExit.Location = new System.Drawing.Point(1222, 1);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(24, 24);
            this.btnExit.TabIndex = 0;
            this.btnExit.TabStop = false;
            this.btnExit.Text = "🗙";
            this.btnExit.UseVisualStyleBackColor = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            this.btnExit.MouseEnter += new System.EventHandler(this.btnExit_MouseEnter);
            this.btnExit.MouseLeave += new System.EventHandler(this.btnExit_MouseLeave);
            // 
            // menuStrip
            // 
            this.menuStrip.BackColor = System.Drawing.Color.Transparent;
            this.menuStrip.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.ediToolStripMenuItem,
            this.renderToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 1);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(180, 24);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importToolStripMenuItem,
            this.exportToolStripMenuItem,
            this.saveProjectToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // importToolStripMenuItem
            // 
            this.importToolStripMenuItem.Name = "importToolStripMenuItem";
            this.importToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.importToolStripMenuItem.Text = "New Project";
            this.importToolStripMenuItem.Click += new System.EventHandler(this.importToolStripMenuItem_Click);
            // 
            // exportToolStripMenuItem
            // 
            this.exportToolStripMenuItem.Name = "exportToolStripMenuItem";
            this.exportToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.exportToolStripMenuItem.Text = "Open Project";
            this.exportToolStripMenuItem.Click += new System.EventHandler(this.exportToolStripMenuItem_Click);
            // 
            // saveProjectToolStripMenuItem
            // 
            this.saveProjectToolStripMenuItem.Name = "saveProjectToolStripMenuItem";
            this.saveProjectToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.saveProjectToolStripMenuItem.Text = "Save Project";
            // 
            // ediToolStripMenuItem
            // 
            this.ediToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.undoToolStripMenuItem,
            this.redoToolStripMenuItem,
            this.toolStripSeparator1,
            this.cutToolStripMenuItem,
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem,
            this.toolStripSeparator2,
            this.deleteToolStripMenuItem});
            this.ediToolStripMenuItem.Name = "ediToolStripMenuItem";
            this.ediToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.ediToolStripMenuItem.Text = "Edit";
            // 
            // undoToolStripMenuItem
            // 
            this.undoToolStripMenuItem.Name = "undoToolStripMenuItem";
            this.undoToolStripMenuItem.RightToLeftAutoMirrorImage = true;
            this.undoToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.undoToolStripMenuItem.Text = "Undo";
            this.undoToolStripMenuItem.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.undoToolStripMenuItem.Click += new System.EventHandler(this.undoToolStripMenuItem_Click);
            // 
            // redoToolStripMenuItem
            // 
            this.redoToolStripMenuItem.Name = "redoToolStripMenuItem";
            this.redoToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.redoToolStripMenuItem.Text = "Redo";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(104, 6);
            // 
            // cutToolStripMenuItem
            // 
            this.cutToolStripMenuItem.Name = "cutToolStripMenuItem";
            this.cutToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.cutToolStripMenuItem.Text = "Cut";
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.pasteToolStripMenuItem.Text = "Paste";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(104, 6);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // renderToolStripMenuItem
            // 
            this.renderToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.createVideoToolStripMenuItem});
            this.renderToolStripMenuItem.Name = "renderToolStripMenuItem";
            this.renderToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.renderToolStripMenuItem.Text = "Export";
            // 
            // createVideoToolStripMenuItem
            // 
            this.createVideoToolStripMenuItem.Name = "createVideoToolStripMenuItem";
            this.createVideoToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.createVideoToolStripMenuItem.Text = "Create Video";
            this.createVideoToolStripMenuItem.Click += new System.EventHandler(this.createVideoToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.infoToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // infoToolStripMenuItem
            // 
            this.infoToolStripMenuItem.Name = "infoToolStripMenuItem";
            this.infoToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.infoToolStripMenuItem.Text = "Info";
            this.infoToolStripMenuItem.Click += new System.EventHandler(this.infoToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.aboutToolStripMenuItem.Text = "About...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // splitVertical
            // 
            this.splitVertical.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitVertical.Location = new System.Drawing.Point(0, 0);
            this.splitVertical.Name = "splitVertical";
            // 
            // splitVertical.Panel1
            // 
            this.splitVertical.Panel1.Controls.Add(this.pnElements);
            // 
            // splitVertical.Panel2
            // 
            this.splitVertical.Panel2.Controls.Add(this.axWindowsMediaPlayer1);
            this.splitVertical.Size = new System.Drawing.Size(1247, 286);
            this.splitVertical.SplitterDistance = 618;
            this.splitVertical.TabIndex = 51;
            // 
            // axWindowsMediaPlayer1
            // 
            this.axWindowsMediaPlayer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axWindowsMediaPlayer1.Enabled = true;
            this.axWindowsMediaPlayer1.Location = new System.Drawing.Point(0, 0);
            this.axWindowsMediaPlayer1.Name = "axWindowsMediaPlayer1";
            this.axWindowsMediaPlayer1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axWindowsMediaPlayer1.OcxState")));
            this.axWindowsMediaPlayer1.Size = new System.Drawing.Size(625, 286);
            this.axWindowsMediaPlayer1.TabIndex = 2;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 428);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1246, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(39, 17);
            this.toolStripStatusLabel1.Text = "Ready";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(200, 16);
            this.toolStripProgressBar1.Visible = false;
            // 
            // splitHorizontal
            // 
            this.splitHorizontal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitHorizontal.Location = new System.Drawing.Point(0, 28);
            this.splitHorizontal.Name = "splitHorizontal";
            this.splitHorizontal.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitHorizontal.Panel1
            // 
            this.splitHorizontal.Panel1.Controls.Add(this.splitVertical);
            // 
            // splitHorizontal.Panel2
            // 
            this.splitHorizontal.Panel2.Controls.Add(this.pnLeftMenu);
            this.splitHorizontal.Panel2.Controls.Add(this.pnVideoEditing);
            this.splitHorizontal.Size = new System.Drawing.Size(1247, 554);
            this.splitHorizontal.SplitterDistance = 286;
            this.splitHorizontal.TabIndex = 52;
            // 
            // pnLeftMenu
            // 
            this.pnLeftMenu.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.pnLeftMenu.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.pnLeftMenu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnLeftMenu.Location = new System.Drawing.Point(0, 0);
            this.pnLeftMenu.Name = "pnLeftMenu";
            this.pnLeftMenu.Size = new System.Drawing.Size(50, 263);
            this.pnLeftMenu.TabIndex = 1;
            this.pnLeftMenu.Paint += new System.Windows.Forms.PaintEventHandler(this.pnMenu_Paint);
            // 
            // fileSystemWatcher
            // 
            this.fileSystemWatcher.EnableRaisingEvents = true;
            this.fileSystemWatcher.SynchronizingObject = this;
            // 
            // VideoEditor
            // 
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.Application;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CancelButton = this.btnExit;
            this.ClientSize = new System.Drawing.Size(1247, 582);
            this.ControlBox = false;
            this.Controls.Add(this.splitHorizontal);
            this.Controls.Add(this.pnTitle);
            this.Controls.Add(this.statusStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip;
            this.Name = "VideoEditor";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ClientSizeChanged += new System.EventHandler(this.FormEditor_ClientSizeChanged);
            this.SizeChanged += new System.EventHandler(this.FormEditor_SizeChanged);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormEditor_KeyDown);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FormEditor_KeyPress);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.FormEditor_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.FormEditor_MouseMove);
            this.Resize += new System.EventHandler(this.FormEditor_Resize);
            this.pnVideoEditing.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbCursor)).EndInit();
            this.pnElements.ResumeLayout(false);
            this.pnTitle.ResumeLayout(false);
            this.pnTitle.PerformLayout();
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.splitVertical.Panel1.ResumeLayout(false);
            this.splitVertical.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitVertical)).EndInit();
            this.splitVertical.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.axWindowsMediaPlayer1)).EndInit();
            this.splitHorizontal.Panel1.ResumeLayout(false);
            this.splitHorizontal.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitHorizontal)).EndInit();
            this.splitHorizontal.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnVideoEditing;
        private AxWMPLib.AxWindowsMediaPlayer axWindowsMediaPlayer1;
        private System.Windows.Forms.PictureBox pbCursor;
        private System.Windows.Forms.Timer cursorTimer;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.HScrollBar hScrollBarZoom;
        private System.Windows.Forms.Panel pnElements;
        private System.Windows.Forms.Panel pnTitle;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ediToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem renderToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnMaximize;
        private System.Windows.Forms.Button btnMinimize;
        private System.Windows.Forms.ToolStripMenuItem importToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem undoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem redoToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem cutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveProjectToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem createVideoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem infoToolStripMenuItem;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.SplitContainer splitVertical;
        private System.Windows.Forms.SplitContainer splitHorizontal;
        private System.IO.FileSystemWatcher fileSystemWatcher;
        private System.Windows.Forms.Panel pnLeftMenu;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}

