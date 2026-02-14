using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WMPLib;
using System.Runtime.InteropServices;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace VideoEditor
{
    public partial class VideoEditor : Form
    {
        protected override bool ShowFocusCues => false;
        private Brush lineColor = Brushes.Black;
        Size visibleSize;
        int INTERVAL_HEIGHT = 30;
        int xIntervalLocation = 0;
        int HEIGHT_LINES = 80;
        int maxNrHeightLines = 8;
        int[] heightIntervals;
        int maxWidth = 8_000;
        int fragment = 6; // 6-36 pixels fragments
        int stage = 0; // 0 - 6 zoomStage
        double[] stageIntervals = new double[6] { 3600, 600, 60, 20, 4, 0.2 }; // all zoom stage intervals
        int MAX_STAGE = 5;
        int previousZoom = 0;
        int maxZoomNrOfPixels;
        int MIN_FRAGMENT = 6;
        int MAX_FRAGMENT = 36;
        int zoom = 0; // current zoom
        double onePixelSec = 30; // 1 pixel is 30 sec
        int currentItem;
        Dictionary<int, Item> items;
        Project currentProject;
        int NEXT_VIDEO_OFFSET = 5;
        int mouseDownButtonX;
        int mouseDownButtonY;
        TrackBar tbVolume;
        bool projectIsPlaying = false;
        
        // Selection, Snapping and Linking
        private List<int> selectedItemIds = new List<int>();
        private bool isSnappingEnabled = true;
        private bool isLinkingEnabled = true;
        private const int SNAP_DISTANCE = 15; // pixels

        // Media Bin
        ListView mediaBinListView;
        ImageList mediaBinImageList;
        List<MediaBinItem> mediaBinItems = new List<MediaBinItem>();

        // Audio Engine
        private IWavePlayer audioOutputDevice;
        private MixingSampleProvider audioMixer;
        private Dictionary<int, AudioFileReader> activeAudioStreams = new Dictionary<int, AudioFileReader>();
        
        // UI Extensions
        private Label lblCurrentTime;
        private Panel pnBlackScreen;
        private bool isPreviewMode = false;
        private TrackBar tbItemVolume;
        private Label lblItemVolumeValue;

        // Render Section
        private double renderStartTime = -1;
        private double renderEndTime = -1;


        void UpdateTotalViewNrOfPixels()
        {
            maxZoomNrOfPixels = visibleSize.Width;
        }
        public VideoEditor()
        {
            InitializeComponent();
            pnVideoEditing.Width = maxWidth;
            pnVideoEditing.Height = maxNrHeightLines * HEIGHT_LINES + INTERVAL_HEIGHT;
            visibleSize = new Size(this.Width - this.DefaultMargin.Horizontal, this.Height - this.DefaultMargin.Vertical - pnVideoEditing.Location.Y);
            UpdateTotalViewNrOfPixels();
            btnMinimize.FlatAppearance.BorderSize = 0;
            btnMaximize.FlatAppearance.BorderSize = 0;
            btnExit.FlatAppearance.BorderSize = 0;
            this.WindowState = FormWindowState.Maximized;
            var fullHeight = maxNrHeightLines * HEIGHT_LINES;
            heightIntervals = new int[maxNrHeightLines + 1];
            for (int i = 0; i <= maxNrHeightLines; i++)
            {
                heightIntervals[i] = HEIGHT_LINES * i;
            }
            axWindowsMediaPlayer1.PlayStateChange += AxWindowsMediaPlayer1_PlayStateChange;
            items = new Dictionary<int, Item>();
            
            // Mouse Wheel & Focus
            pnVideoEditing.MouseEnter += (s, e) => { pnVideoEditing.Focus(); };
            
            // Enable Double Buffering for pnVideoEditing to prevent scroll artifacts
            typeof(Panel).InvokeMember("DoubleBuffered", 
                System.Reflection.BindingFlags.SetProperty | 
                System.Reflection.BindingFlags.Instance | 
                System.Reflection.BindingFlags.NonPublic, 
                null, pnVideoEditing, new object[] { true });

            // Ensure we redraw when scrolling (key for keeping grid lines synced)
            pnVideoEditing.Scroll += (s, e) => pnVideoEditing.Invalidate();

            currentProject = new Project();
            lblProjectName.Text = "Current Project: " + currentProject.projectName;
            lblProjectName.Location = new Point(this.Width / 2 - lblProjectName.Width / 2, lblProjectName.Location.Y);

            this.pnVideoEditing.MouseWheel += pnVideoEditing_MouseWheel;

            // Volume Control
            Panel pnVideoContainer = new Panel();
            pnVideoContainer.Dock = DockStyle.Fill;
            splitVertical.Panel2.Controls.Add(pnVideoContainer);

            tbVolume = new TrackBar();
            tbVolume.Minimum = 0;
            tbVolume.Maximum = 100;
            tbVolume.Value = axWindowsMediaPlayer1.settings.volume;
            tbVolume.TickFrequency = 10;
            tbVolume.Height = 30;
            tbVolume.Dock = DockStyle.Bottom;
            tbVolume.Dock = DockStyle.Bottom;
            tbVolume.ValueChanged += (s, e) => {
                axWindowsMediaPlayer1.settings.volume = tbVolume.Value;
                if (audioOutputDevice != null)
                {
                    audioOutputDevice.Volume = tbVolume.Value / 100f;
                }
            };
            pnVideoContainer.Controls.Add(tbVolume);
            
            axWindowsMediaPlayer1.Parent = pnVideoContainer;
            axWindowsMediaPlayer1.Dock = DockStyle.Fill;
            axWindowsMediaPlayer1.BringToFront();

            // Scrolling sync
            pnVideoEditing.Scroll += (s, e) => {
                pnLeftMenu.Invalidate();
            };

            // Setup the Media Bin in the left panel
            SetupMediaBin();

            // Setup audio engine
            InitAudioEngine();

            // Setup extra UI elements
            SetupExtraUI();

            // Setup toolbar buttons
            SetupToolbar();
        }

        private void SetupExtraUI()
        {
            // Current Time Label
            lblCurrentTime = new Label();
            lblCurrentTime.ForeColor = Color.White;
            lblCurrentTime.BackColor = Color.Transparent;
            lblCurrentTime.Font = new Font("Consolas", 14, FontStyle.Bold);
            lblCurrentTime.Text = "00:00:00:00";
            lblCurrentTime.AutoSize = true;
            lblCurrentTime.Location = new Point(10, 10);
            // Add to the container that holds the player. 
            // tbVolume is in a pnVideoContainer creating in constructor, but local variable.
            // We can add to splitVertical.Panel2 and BringToFront
            splitVertical.Panel2.Controls.Add(lblCurrentTime);
            lblCurrentTime.BringToFront();

            // Black Screen Panel
            pnBlackScreen = new Panel();
            pnBlackScreen.BackColor = Color.Black;
            pnBlackScreen.Dock = DockStyle.Fill;
            pnBlackScreen.Visible = false;
            axWindowsMediaPlayer1.Parent.Controls.Add(pnBlackScreen);
            pnBlackScreen.SendToBack();

            // Item Volume Control in Left Menu
            Label lblVol = new Label();
            lblVol.Text = "Vol:";
            lblVol.Font = new Font("Segoe UI", 8);
            lblVol.Location = new Point(54, 8);
            lblVol.AutoSize = true;
            pnLeftMenu.Controls.Add(lblVol);

            tbItemVolume = new TrackBar();
            tbItemVolume.Minimum = 0;
            tbItemVolume.Maximum = 125;
            tbItemVolume.Value = 100;
            tbItemVolume.TickStyle = TickStyle.None;
            tbItemVolume.Size = new Size(80, 20);
            tbItemVolume.Location = new Point(80, 6);
            tbItemVolume.ValueChanged += TbItemVolume_ValueChanged;
            pnLeftMenu.Controls.Add(tbItemVolume);

            lblItemVolumeValue = new Label();
            lblItemVolumeValue.Text = "100%";
            lblItemVolumeValue.Font = new Font("Segoe UI", 7);
            lblItemVolumeValue.Location = new Point(160, 8);
            lblItemVolumeValue.AutoSize = true;
            pnLeftMenu.Controls.Add(lblItemVolumeValue);
        }

        private void TbItemVolume_ValueChanged(object sender, EventArgs e)
        {
            float vol = tbItemVolume.Value / 100f;
            lblItemVolumeValue.Text = tbItemVolume.Value + "%";
            
            if (currentItem > 0 && items.ContainsKey(currentItem))
            {
                var item = items[currentItem];
                if (selectedItemIds.Count > 0)
                {
                     foreach(var id in selectedItemIds)
                     {
                         if(items.ContainsKey(id)) UpdateItemVolume(items[id], vol);
                     }
                }
                else
                {
                     UpdateItemVolume(item, vol);
                }
            }
        }

        private void InitAudioEngine()
        {
            try
            {
                audioOutputDevice = new WaveOutEvent();
                // Create a mixer with stereo output, 44.1kHz
                audioMixer = new MixingSampleProvider(WaveFormat.CreateIeeeFloatWaveFormat(44100, 2));
                audioMixer.ReadFully = true; // Keep mixer alive even if no inputs
                audioOutputDevice.Init(audioMixer);
                audioOutputDevice.Play();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Audio Init Failed: " + ex.Message);
            }
        }

        private void UpdateAudioPlayback(double currentTime)
        {
            // Simple approach: When playing, constantly check needed streams
            // Ideally should be event driven but polling in timer works for simple editing
            if (!projectIsPlaying) 
            {
                 StopAudioPlayback();
                 return;
            }

            // Identify items that SHOULD be playing
            // Only Audio Items (type == .wav/mp3 etc) OR Video items if we want to support mixing video audio distinct from WMP.
            // For now, we mix ONLY items where we explicitly want NAudio to handle it.
            // NOTE: WMP handles the main video track audio. If we have overlapping audios, we use NAudio.
            // ALSO: If we are in user removed video mode but kept audio, that logic uses items with fileType Audio.
            
            var audioItems = items.Values.Where(x => x.fileType == FileType.Audio).ToList();
            
            foreach (var item in audioItems)
            {
                int key = items.FirstOrDefault(x => x.Value == item).Key;
                
                // Check if intersects
                if (currentTime >= item.startPoint && currentTime < item.endPoint)
                {
                    // Should be playing
                    if (!activeAudioStreams.ContainsKey(key))
                    {
                        try
                        {
                            var reader = new AudioFileReader(item.path);
                            reader.Volume = item.Volume;
                            
                            // Seek
                            double offset = currentTime - item.startPoint;
                            if (offset < 0) offset = 0;
                            reader.CurrentTime = TimeSpan.FromSeconds(offset);

                            // Resample to match mixer (44100Hz) and ensure stereo to fix pitch issues
                            ISampleProvider resampledInput = new WdlResamplingSampleProvider(reader, 44100);
                            if (resampledInput.WaveFormat.Channels == 1)
                            {
                                resampledInput = new MonoToStereoSampleProvider(resampledInput);
                            }
                            
                            activeAudioStreams.Add(key, reader);
                            audioMixer.AddMixerInput(resampledInput); 
                        }
                        catch (Exception ex) { Console.WriteLine("Audio Play Error: " + ex.Message); }
                    }
                }
                else
                {
                    // Should NOT be playing
                    if (activeAudioStreams.ContainsKey(key))
                    {
                        var reader = activeAudioStreams[key];
                        audioMixer.RemoveMixerInput(reader);
                        reader.Dispose();
                        activeAudioStreams.Remove(key);
                    }
                }
            }
        }

        private void StopAudioPlayback()
        {
            foreach (var reader in activeAudioStreams.Values)
            {
                audioMixer.RemoveMixerInput(reader);
                reader.Dispose();
            }
            activeAudioStreams.Clear();
        }

        private void SetupToolbar()
        {
            pnLeftMenu.Controls.Clear();

            // Snapping button
            Button btnSnapping = new Button();
            btnSnapping.Text = "🧲";
            btnSnapping.Font = new Font("Segoe UI", 10);
            btnSnapping.Size = new Size(22, 22);
            btnSnapping.Location = new Point(2, 4);
            btnSnapping.FlatStyle = FlatStyle.Flat;
            btnSnapping.FlatAppearance.BorderSize = isSnappingEnabled ? 2 : 1;
            btnSnapping.FlatAppearance.BorderColor = isSnappingEnabled ? Color.Blue : Color.Gray;
            btnSnapping.BackColor = isSnappingEnabled ? Color.LightBlue : Color.White;
            btnSnapping.Click += (s, e) => {
                isSnappingEnabled = !isSnappingEnabled;
                btnSnapping.BackColor = isSnappingEnabled ? Color.LightBlue : Color.White;
                btnSnapping.FlatAppearance.BorderSize = isSnappingEnabled ? 2 : 1;
                btnSnapping.FlatAppearance.BorderColor = isSnappingEnabled ? Color.Blue : Color.Gray;
            };
            ToolTip ttSnap = new ToolTip();
            ttSnap.SetToolTip(btnSnapping, "Toggle Snapping (Align clips automatically)");
            pnLeftMenu.Controls.Add(btnSnapping);

            // Linking button
            Button btnLinking = new Button();
            btnLinking.Text = "🔗";
            btnLinking.Font = new Font("Segoe UI", 10);
            btnLinking.Size = new Size(22, 22);
            btnLinking.Location = new Point(26, 4);
            btnLinking.FlatStyle = FlatStyle.Flat;
            btnLinking.FlatAppearance.BorderSize = isLinkingEnabled ? 2 : 1;
            btnLinking.FlatAppearance.BorderColor = isLinkingEnabled ? Color.Blue : Color.Gray;
            btnLinking.BackColor = isLinkingEnabled ? Color.LightBlue : Color.White;
            btnLinking.Click += (s, e) => {
                isLinkingEnabled = !isLinkingEnabled;
                btnLinking.BackColor = isLinkingEnabled ? Color.LightBlue : Color.White;
                btnLinking.FlatAppearance.BorderSize = isLinkingEnabled ? 2 : 1;
                btnLinking.FlatAppearance.BorderColor = isLinkingEnabled ? Color.Blue : Color.Gray;
            };
            ToolTip ttLink = new ToolTip();
            ttLink.SetToolTip(btnLinking, "Toggle Clip Linking (Move Audio+Video together)");
            pnLeftMenu.Controls.Add(btnLinking);
        }

        /// <summary>
        /// Media Bin item to store file metadata
        /// </summary>
        class MediaBinItem
        {
            public string Path { get; set; }
            public string FileName { get; set; }
            public double Duration { get; set; }
            public string Extension { get; set; }
            public FileType FileType { get; set; }
        }

        void SetupMediaBin()
        {
            // Clear existing controls except the zoom scrollbar
            var zoomBar = hScrollBarZoom;
            pnElements.Controls.Clear();

            // Header label
            Label lblHeader = new Label();
            lblHeader.Text = "📁 Media Bin";
            lblHeader.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblHeader.ForeColor = Color.FromArgb(60, 60, 60);
            lblHeader.BackColor = Color.FromArgb(235, 235, 240);
            lblHeader.Dock = DockStyle.Top;
            lblHeader.Height = 28;
            lblHeader.TextAlign = ContentAlignment.MiddleLeft;
            lblHeader.Padding = new Padding(6, 0, 0, 0);
            pnElements.Controls.Add(lblHeader);

            // Zoom scrollbar at bottom
            zoomBar.Dock = DockStyle.Bottom;
            pnElements.Controls.Add(zoomBar);

            // Image list for thumbnails
            mediaBinImageList = new ImageList();
            mediaBinImageList.ImageSize = new Size(64, 48);
            mediaBinImageList.ColorDepth = ColorDepth.Depth32Bit;

            // Add default type icons
            // Video icon
            Bitmap videoIcon = new Bitmap(64, 48);
            using (Graphics g = Graphics.FromImage(videoIcon))
            {
                g.Clear(Color.FromArgb(40, 40, 50));
                g.DrawString("🎬", new Font("Segoe UI", 18), Brushes.White, 14, 6);
            }
            mediaBinImageList.Images.Add("video", videoIcon);

            // Audio icon
            Bitmap audioIcon = new Bitmap(64, 48);
            using (Graphics g = Graphics.FromImage(audioIcon))
            {
                g.Clear(Color.FromArgb(30, 60, 30));
                g.DrawString("🎵", new Font("Segoe UI", 18), Brushes.White, 14, 6);
            }
            mediaBinImageList.Images.Add("audio", audioIcon);

            // Image icon
            Bitmap imageIcon = new Bitmap(64, 48);
            using (Graphics g = Graphics.FromImage(imageIcon))
            {
                g.Clear(Color.FromArgb(50, 30, 60));
                g.DrawString("🖼", new Font("Segoe UI", 18), Brushes.White, 14, 6);
            }
            mediaBinImageList.Images.Add("image", imageIcon);

            // ListView
            mediaBinListView = new ListView();
            mediaBinListView.Dock = DockStyle.Fill;
            mediaBinListView.View = View.Details;
            mediaBinListView.FullRowSelect = true;
            mediaBinListView.GridLines = true;
            mediaBinListView.SmallImageList = mediaBinImageList;
            mediaBinListView.LargeImageList = mediaBinImageList;
            mediaBinListView.BackColor = Color.FromArgb(245, 245, 248);
            mediaBinListView.Font = new Font("Segoe UI", 8.5f);
            mediaBinListView.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            mediaBinListView.BorderStyle = BorderStyle.None;

            // Columns
            mediaBinListView.Columns.Add("Name", 200);
            mediaBinListView.Columns.Add("Duration", 65);
            mediaBinListView.Columns.Add("Type", 55);

            // Drag-drop to import files into the bin
            mediaBinListView.AllowDrop = true;
            mediaBinListView.DragEnter += MediaBin_DragEnter;
            mediaBinListView.DragDrop += MediaBin_DragDrop;

            // Double-click to add to timeline
            mediaBinListView.MouseDoubleClick += MediaBin_DoubleClick;

            // Drag from bin to timeline
            mediaBinListView.ItemDrag += MediaBin_ItemDrag;

            // Context menu for bin items
            ContextMenu binContextMenu = new ContextMenu();
            binContextMenu.MenuItems.Add("Play", (s, e) => {
                if (mediaBinListView.SelectedIndices.Count > 0)
                {
                    var binItem = mediaBinItems[mediaBinListView.SelectedIndices[0]];
                    projectIsPlaying = false;
                    isPreviewMode = true;
                    axWindowsMediaPlayer1.uiMode = "mini";
                    axWindowsMediaPlayer1.URL = binItem.Path;
                    axWindowsMediaPlayer1.Ctlcontrols.play();
                }
            });
            binContextMenu.MenuItems.Add("Add to Timeline", (s, e) => {
                if (mediaBinListView.SelectedIndices.Count > 0)
                {
                    var binItem = mediaBinItems[mediaBinListView.SelectedIndices[0]];
                    AddNewFile(binItem.Path, binItem.FileName, binItem.Duration, binItem.Extension);
                }
            });
            binContextMenu.MenuItems.Add("Remove from Bin", (s, e) => {
                if (mediaBinListView.SelectedIndices.Count > 0)
                {
                    int idx = mediaBinListView.SelectedIndices[0];
                    mediaBinItems.RemoveAt(idx);
                    mediaBinListView.Items.RemoveAt(idx);
                }
            });
            mediaBinListView.ContextMenu = binContextMenu;

            pnElements.Controls.Add(mediaBinListView);

            // Ensure correct z-order: header on top, zoom at bottom, list in middle
            mediaBinListView.BringToFront();
            lblHeader.BringToFront();
        }

        private void MediaBin_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void MediaBin_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null) return;

            foreach (var file in files)
            {
                AddToMediaBin(file);
            }
        }

        private void AddToMediaBin(string filePath)
        {
            // Check if already in bin
            if (mediaBinItems.Any(x => x.Path == filePath)) return;

            string extension = Path.GetExtension(filePath).ToLower();
            string fileName = Path.GetFileName(filePath);
            FileType fileType = GetMediaBinFileType(extension);

            if (fileType == FileType.Unknown) return;

            double duration = Duration(filePath);

            var binItem = new MediaBinItem
            {
                Path = filePath,
                FileName = fileName,
                Duration = duration,
                Extension = extension,
                FileType = fileType
            };
            mediaBinItems.Add(binItem);

            // Generate thumbnail for video files
            string imageKey;
            if (fileType == FileType.Video)
            {
                imageKey = "video";
                try
                {
                    var projectPath = currentProject.projectsPath + currentProject.projectName + @"\";
                    Image thumb = VideoProcessing.GetVideoTumbnail(filePath, 1, projectPath);
                    if (thumb != null)
                    {
                        // Resize to fit image list
                        Bitmap resized = new Bitmap(64, 48);
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.DrawImage(thumb, 0, 0, 64, 48);
                        }
                        string key = "thumb_" + mediaBinItems.Count;
                        mediaBinImageList.Images.Add(key, resized);
                        imageKey = key;
                    }
                }
                catch { }
            }
            else if (fileType == FileType.Audio)
            {
                imageKey = "audio";
            }
            else
            {
                imageKey = "image";
                try
                {
                    Image img = Image.FromFile(filePath);
                    Bitmap resized = new Bitmap(64, 48);
                    using (Graphics g = Graphics.FromImage(resized))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(img, 0, 0, 64, 48);
                    }
                    string key = "thumb_" + mediaBinItems.Count;
                    mediaBinImageList.Images.Add(key, resized);
                    imageKey = key;
                    img.Dispose();
                }
                catch { }
            }

            ListViewItem lvi = new ListViewItem(fileName, imageKey);
            lvi.SubItems.Add(DurationReadeble(duration));
            lvi.SubItems.Add(fileType.ToString());
            lvi.Tag = binItem;
            mediaBinListView.Items.Add(lvi);
        }

        private void MediaBin_DoubleClick(object sender, MouseEventArgs e)
        {
            if (mediaBinListView.SelectedIndices.Count > 0)
            {
                var binItem = mediaBinItems[mediaBinListView.SelectedIndices[0]];
                AddNewFile(binItem.Path, binItem.FileName, binItem.Duration, binItem.Extension);
            }
        }

        private void MediaBin_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (mediaBinListView.SelectedIndices.Count > 0)
            {
                var binItem = mediaBinItems[mediaBinListView.SelectedIndices[0]];
                // Use a custom format so the timeline can recognize it
                DataObject data = new DataObject();
                data.SetData("MediaBinItem", binItem);
                mediaBinListView.DoDragDrop(data, DragDropEffects.Copy);
            }
        }

        private FileType GetMediaBinFileType(string extension)
        {
            if (Constants.SUPPORTED_VIDEO_FORMATS.Contains(extension)) return FileType.Video;
            if (Constants.SUPPORTED_AUDIO_FORMATS.Contains(extension)) return FileType.Audio;
            if (Constants.SUPPORTED_IMAGE_FORMATS.Contains(extension)) return FileType.Image;
            return FileType.Unknown;
        }

        private void pnVideoEditing_MouseWheel(object sender, MouseEventArgs e)
        {
            if (ModifierKeys == Keys.Shift)
            {
                // Horizontal scroll logic - check Shift FIRST so it always works regardless of mouse position
                int scrollAmount = Math.Abs(e.Delta) > 0 ? Math.Abs(e.Delta) : 50;
                Point currentPos = pnVideoEditing.AutoScrollPosition;
                if (e.Delta > 0)
                {
                    pnVideoEditing.AutoScrollPosition = new Point(Math.Max(0, Math.Abs(currentPos.X) - scrollAmount), Math.Abs(currentPos.Y));
                }
                else
                {
                    pnVideoEditing.AutoScrollPosition = new Point(Math.Abs(currentPos.X) + scrollAmount, Math.Abs(currentPos.Y));
                }
                // Suppress default ScrollableControl horizontal scroll behavior
                if (e is HandledMouseEventArgs he) he.Handled = true;
            }
            else if (ModifierKeys == Keys.Control || e.Y < INTERVAL_HEIGHT)
            {
                // Zoom logic
                if (e.Delta > 0)
                {
                    if (hScrollBarZoom.Value < hScrollBarZoom.Maximum)
                        hScrollBarZoom.Value++;
                }
                else
                {
                    if (hScrollBarZoom.Value > hScrollBarZoom.Minimum)
                        hScrollBarZoom.Value--;
                }
                // Suppress default scroll behavior during zoom
                if (e is HandledMouseEventArgs he) he.Handled = true;
            }
            else
            {
                // Disable vertical scroll in the timeline - suppress default behavior
                if (e is HandledMouseEventArgs he) he.Handled = true;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            pbCursor.Location = new Point(0, 0);
            pbCursor.Height = HEIGHT_LINES * (maxNrHeightLines - 1);
            pbCursor.BringToFront();
            pnVideoEditing.AutoScrollPosition = new Point(0, 0);
            cursorTimer.Start();
            pnVideoEditing.Update();
            hScrollBarZoom.Value = MAX_STAGE / 2;
            hScrollBarZoom.BringToFront();
        }
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr one, int two, int three, int four);
        private void AxWindowsMediaPlayer1_PlayStateChange(object sender, AxWMPLib._WMPOCXEvents_PlayStateChangeEvent e)
        {
            // Transitions and gaps are now handled in timer1_Tick via projectIsPlaying logic
        }
        void DrawLeftSideLines(Graphics g)
        {
            if (g == null) return;
            // Align with vertical scroll of the main panel
            g.TranslateTransform(0, pnVideoEditing.AutoScrollPosition.Y);
            // horizontal line
            for (int i = 0; i <= maxNrHeightLines; i++)
            {
                int y = i * HEIGHT_LINES + INTERVAL_HEIGHT;
                g.DrawLine(new Pen(lineColor), new Point(0, y),
                                                         new Point(pnLeftMenu.Width, y));

            }
            // Track type labels
            string[] trackLabels = { "Images", "Video", "Audio 1", "Audio 2" };
            using (Font labelFont = new Font("Segoe UI", 8, FontStyle.Bold))
            {
                for (int i = 0; i < Math.Min(trackLabels.Length, maxNrHeightLines); i++)
                {
                    int y = i * HEIGHT_LINES + INTERVAL_HEIGHT + (HEIGHT_LINES / 2) - 8;
                    var sf = new StringFormat() { Alignment = StringAlignment.Center };
                    g.DrawString(trackLabels[i], labelFont, lineColor, pnLeftMenu.Width / 2, y, sf);
                }
            }
        }
        void DrawVideoEditingLines(Graphics g)
        {
            if (g == null) return;

            // Apply scroll transform to align drawing with scrolled content
            g.TranslateTransform(pnVideoEditing.AutoScrollPosition.X, pnVideoEditing.AutoScrollPosition.Y);

            // Calculate total width based on items
            int maxItemX = 0;
            if (items.Count > 0)
            {
                 maxItemX = items.Values.Max(x => (int)((x.startPoint + x.duration) / onePixelSec));
            }
            // Use virtual width 
            int nrTotalPixels = Math.Max(pnVideoEditing.Width, pnVideoEditing.AutoScrollMinSize.Width);
            
            // Draw ticks
            for (int i = 0; i < nrTotalPixels; i += fragment)
            {
                g.DrawLine(new Pen(lineColor), new Point(i, INTERVAL_HEIGHT),
                                                        new Point(i, INTERVAL_HEIGHT - 4));
                // Minor ticks every 5 fragments
                if (i % (fragment * 5) == 0)
                {
                    g.DrawLine(new Pen(lineColor), new Point(i, INTERVAL_HEIGHT),
                                        new Point(i, INTERVAL_HEIGHT - 8));
                }
            }

            // Draw labels based on pixel distance to avoid overlap
            int labelStep = 100; // minimum pixels between labels
            int timeStepPixels = (int)Math.Ceiling((double)labelStep / fragment) * fragment;
            for (int i = 0; i < nrTotalPixels; i += timeStepPixels)
            {
                // Show frames/milliseconds for precision
                string timeStr = TimeSpan.FromSeconds(onePixelSec * i).ToString(@"hh\:mm\:ss\:ff");
                g.DrawString(timeStr, SystemFonts.DefaultFont, lineColor, i, 6);
            }

            // Calculate Virtual Height of the timeline
            int virtualHeight = Math.Max(pnVideoEditing.Height, HEIGHT_LINES * maxNrHeightLines + INTERVAL_HEIGHT);

            // Draw Render Section
            if (renderStartTime >= 0 && renderEndTime > renderStartTime)
            {
                int startX = (int)(renderStartTime / onePixelSec);
                int endX = (int)(renderEndTime / onePixelSec);
                int width = endX - startX;
                
                // Draw shaded region
                using (Brush brush = new SolidBrush(Color.FromArgb(50, 0, 0, 255)))
                {
                    g.FillRectangle(brush, startX, INTERVAL_HEIGHT, width, virtualHeight - INTERVAL_HEIGHT);
                }
                
                // Draw markers
                g.DrawLine(Pens.Blue, startX, 0, startX, virtualHeight);
                g.DrawLine(Pens.Blue, endX, 0, endX, virtualHeight);
                
                // Label
                g.DrawString("Render Start", SystemFonts.DefaultFont, Brushes.Blue, startX + 2, 2);
                g.DrawString("Render End", SystemFonts.DefaultFont, Brushes.Blue, endX - 60, 2);
            }

            // horizontal grid lines
            for (int i = 0; i * HEIGHT_LINES + INTERVAL_HEIGHT <= virtualHeight; i++)
            {
                g.DrawLine(new Pen(lineColor), new Point(0, i * HEIGHT_LINES + INTERVAL_HEIGHT),
                                                        new Point(nrTotalPixels, i * HEIGHT_LINES + INTERVAL_HEIGHT));
            }
        }
        /// <summary>
        /// Drag & Drop
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pnVideoEditing_Paint(object sender, PaintEventArgs e)
        {
            DrawVideoEditingLines(e.Graphics);
        }
        private void pnMenu_Paint(object sender, PaintEventArgs e)
        {
            DrawLeftSideLines(e.Graphics);
        }

        void RefreshItemsPositionAndSize()
        {
            int maxLocationX = 0;
            foreach (var key in items.Keys)
            {
                Button crtButton = items[key].button;
                crtButton.Location = new Point((int)Math.Round(items[key].startPoint / onePixelSec), items[key].grid * HEIGHT_LINES + INTERVAL_HEIGHT);
                crtButton.Width = (int)Math.Round(items[key].duration / onePixelSec);
                crtButton.Height = HEIGHT_LINES;
                crtButton.FlatStyle = FlatStyle.Flat;
                crtButton.FlatAppearance.BorderSize = 0;

                int endTime = crtButton.Location.X + crtButton.Width;
                if (endTime > maxLocationX) maxLocationX = endTime;
            }
            
            // Adjust panel layout for scrolling
            
            // Explicitly set the virtual canvas size using AutoScrollMinSize.
            // Only set width for horizontal scroll - no vertical scroll needed.
            int virtualWidth = Math.Max(visibleSize.Width, maxLocationX + 2000);
            int virtualHeight = HEIGHT_LINES * maxNrHeightLines + INTERVAL_HEIGHT;
            
            pnVideoEditing.AutoScrollMinSize = new Size(virtualWidth, 0);
            
            // Hide vertical scrollbar
            pnVideoEditing.VerticalScroll.Visible = false;
            pnVideoEditing.VerticalScroll.Enabled = false;
            
            // Ensure Anchor is set to fill parent visual area, so scrollbars appear on the edges of the PARENT/Panel
            pnVideoEditing.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // Update cursor height to match content height
            if (pbCursor != null && pbCursor.Height != virtualHeight)
            {
                pbCursor.Height = virtualHeight;
            }
        }
        // Removed duplicate PnVideoEditing_MouseWheel handler
        // All mouse wheel handling is now consolidated in pnVideoEditing_MouseWheel

        private void pnVideoEditing_DragDrop(object sender, DragEventArgs e)
        {
            // Check if dragged from Media Bin
            if (e.Data.GetDataPresent("MediaBinItem"))
            {
                var binItem = (MediaBinItem)e.Data.GetData("MediaBinItem");
                AddNewFile(binItem.Path, binItem.FileName, binItem.Duration, binItem.Extension);
                return;
            }

            // File drop from Explorer
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            string[] formats = e.Data.GetFormats();
            if (files?.Count() > 0)
            {
                foreach (var file in files)
                {
                    double duration = Duration(file);
                    string extension = Path.GetExtension(file);
                    string fileName = Path.GetFileName(file);
                    // Also add to Media Bin for reference
                    AddToMediaBin(file);
                    AddNewFile(file, fileName, duration, extension);
                }

            }
            else
            {
                MessageBox.Show("No files added!");
            }


        }
        void UpdateVideoEditingMaxWidth(int maxPosition)
        {
            int position = maxPosition + (int)(0.3 * maxPosition);
            if (visibleSize.Width < position)
            {
                pnVideoEditing.Width = maxPosition + (int)(0.3 * maxPosition);
            }
        }

        private void AddVideoImage(string fileName, string projectPath, string path, int duration, Button crtButton)
        {
            // Generate thumbnails for the video button
            List<Image> alignedImages = new List<Image>();
            float i = 0;
            float timeStep = Math.Max(1.0f, (float)(duration / 20.0f)); // ~20 thumbnails max

            while (i < duration)
            {
                Image crtImg = VideoProcessing.GetVideoTumbnail(path, i, projectPath);
                if (crtImg == null)
                {
                    i += timeStep;
                    continue;
                }
                
                alignedImages.Add(crtImg);
                i += timeStep;
                if (i <= 0) break;
            }

            if (alignedImages.Count > 0)
            {
                int concImgWidth = alignedImages.Sum(x => x.Width);
                int concMaxHeight = alignedImages.Max(x => x.Height);
                var bitmap = new Bitmap(concImgWidth, concMaxHeight);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    int xpos = 0;
                    foreach (var image in alignedImages)
                    {
                        g.DrawImage(image, xpos, 0);
                        xpos += image.Width;
                    }
                }
                crtButton.FlatStyle = FlatStyle.Flat;
                crtButton.FlatAppearance.BorderSize = 0;
                crtButton.BackgroundImageLayout = ImageLayout.Stretch;
                crtButton.BackgroundImage = bitmap;
            }
        }
        
        /// <summary>
        /// Creates a separate audio waveform track for a video file
        /// </summary>
        private void AddVideoAudioTrack(string fileName, string path, double duration, double startPos, int videoGrid, int videoKey)
        {
            var projectPath = currentProject.projectsPath + currentProject.projectName + @"\";
            var wavFileName = VideoProcessing.FileNameWithoutExtension(fileName) + ".wav";
            var outputWav = projectPath + wavFileName;
            VideoProcessing.ConvertFileToWAV(path, (int)duration, outputWav);
            
            Image waveImage = null;
            if (File.Exists(outputWav))
            {
                waveImage = VideoProcessing.CreateWaveImage(outputWav, (int)duration);
            }
            
            if (waveImage == null) return;

            int audioGrid = videoGrid + 1; // Audio track right below video
            int audioKey = items.Keys.Max() + 1;

            Button audioButton = new Button();
            audioButton.FlatStyle = FlatStyle.Flat;
            audioButton.FlatAppearance.BorderSize = 0;
            audioButton.Text = wavFileName + " : " + DurationReadeble(duration);
            audioButton.Name = path + "," + duration + "," + audioKey;
            audioButton.Cursor = Cursors.Hand;
            audioButton.Click += Item_Button_Click;
            audioButton.MouseUp += Item_Button_MouseUp;
            audioButton.MouseDown += Item_Button_MouseDown;
            audioButton.MouseMove += Item_Button_MouseMove;
            audioButton.Size = new Size((int)Math.Round(duration / onePixelSec), HEIGHT_LINES);
            audioButton.Height = HEIGHT_LINES;
            audioButton.BackgroundImageLayout = ImageLayout.Stretch;
            audioButton.BackgroundImage = waveImage;
            pnVideoEditing.Controls.Add(audioButton);
            audioButton.Location = new Point((int)Math.Round(startPos / onePixelSec), audioGrid * HEIGHT_LINES + INTERVAL_HEIGHT);

            Item audioItem = new Item
            {
                path = outputWav,
                fileName = wavFileName,
                fileType = FileType.Audio,
                type = ".wav",
                grid = audioGrid,
                duration = duration,
                startVideo = 0,
                endVideo = duration,
                button = audioButton,
                startPoint = startPos
            };
            audioItem.endPoint = audioItem.startPoint + audioItem.duration;
            items.Add(audioKey, audioItem);

            // Link them
            items[videoKey].linkedItemId = audioKey;
            items[audioKey].linkedItemId = videoKey;
        }
        private void AddNewFile(string path, string fileName, double duration, string extension)
        {
            Cursor = Cursors.WaitCursor;
            FileType fileType;
            if (Constants.SUPPORTED_VIDEO_FORMATS.Contains(extension))
            {
                fileType = FileType.Video;
            }
            else if (Constants.SUPPORTED_AUDIO_FORMATS.Contains(extension))
            {
                fileType = FileType.Audio;
            }
            else if (Constants.SUPPORTED_IMAGE_FORMATS.Contains(extension))
            {
                fileType = FileType.Image;
            }
            else
            {
                MessageBox.Show("Supported video extensions: " +
               string.Join(", ", Constants.SUPPORTED_VIDEO_FORMATS) + ".\nSupported audio extensions:" +
               string.Join(", ", Constants.SUPPORTED_AUDIO_FORMATS) + ".\nSupported image extensions: \n" +
               string.Join(", ", Constants.SUPPORTED_IMAGE_FORMATS) + ".", "File extension not supported!");
                return;
            }
            int key = 0;
            double startPos = 0;
            int grid;
            
            // Video goes on grid 1, its audio on grid 2, standalone audio on grid 3
            if (fileType == FileType.Video)
            {
                grid = 1;
                // Find the max endPoint among video items on grid 1
                var videoItems = items.Values.Where(x => x.grid == 1);
                if (videoItems.Any()) startPos = videoItems.Max(x => x.endPoint);
            }
            else if (fileType == FileType.Audio)
            {
                grid = 3;
                var audioItems = items.Values.Where(x => x.grid == 3);
                if (audioItems.Any()) startPos = audioItems.Max(x => x.endPoint);
            }
            else
            {
                grid = 0; // Images on grid 0
                var imageItems = items.Values.Where(x => x.grid == 0);
                if (imageItems.Any()) startPos = imageItems.Max(x => x.endPoint);
            }

            if (items.Count > 0)
            {
                key = items.Keys.Max() + 1;
            }
            else
            {
                key = 1;
            }
            Button crtButton = new Button();
            crtButton.FlatStyle = FlatStyle.Flat;
            crtButton.FlatAppearance.BorderSize = 0;
            crtButton.Text = fileName + " : " + DurationReadeble(duration);
            crtButton.Name = path + "," + duration + "," + key;
            crtButton.Cursor = Cursors.Hand;
            crtButton.Click += Item_Button_Click;
            crtButton.MouseUp += Item_Button_MouseUp;
            crtButton.MouseDown += Item_Button_MouseDown;
            crtButton.MouseMove += Item_Button_MouseMove;
            crtButton.Size = new Size((int)Math.Round(duration / onePixelSec), HEIGHT_LINES);
            crtButton.Location = new Point((int)Math.Round(startPos / onePixelSec), grid * HEIGHT_LINES + INTERVAL_HEIGHT);
            UpdateVideoEditingMaxWidth(crtButton.Location.X + crtButton.Width);
            pnVideoEditing.Controls.Add(crtButton);
            Item item = new Item
            {
                path = path,
                fileName = fileName,
                fileType = fileType,
                type = extension,
                grid = grid,
                duration = duration,
                startVideo = 0,
                endVideo = duration,
                button = crtButton,
                startPoint = (double)crtButton.Location.X * onePixelSec
            };
            item.endPoint = item.startPoint + item.duration;
            
            // Video: thumbnails on button + auto-extract audio to separate track below
            if (fileType == FileType.Video)
            {
                var projectPath = currentProject.projectsPath + currentProject.projectName + @"\";
                AddVideoImage(fileName, projectPath, path, (int)duration, crtButton);
                items.Add(key, item);
                // Auto-create audio waveform track on grid 2 (below video)
                AddVideoAudioTrack(fileName, path, duration, item.startPoint, grid, key);
            }
            // Standalone audio: waveform on button
            else if (fileType == FileType.Audio)
            {
                if (extension != ".wav")
                {
                    var wavFileName = VideoProcessing.FileNameWithoutExtension(fileName + ".wav");
                    var output = currentProject.projectsPath + currentProject.projectName + @"\" + wavFileName;
                    VideoProcessing.ConvertFileToWAV(path, (int)duration, output);
                    Image wave = VideoProcessing.CreateWaveImage(output, (int)duration);
                    crtButton.BackgroundImageLayout = ImageLayout.Stretch;
                    crtButton.BackgroundImage = wave;
                }
                else if (extension == ".wav")
                {
                    Image wave = VideoProcessing.CreateWaveImage(path, (int)duration);
                    crtButton.BackgroundImageLayout = ImageLayout.Stretch;
                    crtButton.BackgroundImage = wave;
                }
                items.Add(key, item);
            }
            // Images
            else
            {
                items.Add(key, item);
            }
            
            LoadItem(key);
            currentItem = key;
            
            // Exit preview mode if we were in it
            if (isPreviewMode)
            {
                isPreviewMode = false;
                axWindowsMediaPlayer1.uiMode = "none";
            }
            
            pnVideoEditing.Update();
            Cursor = Cursors.Default;
        }

        private void Item_Button_MouseMove(object sender, MouseEventArgs e)
        {
            Button crtButton = sender as Button;
            if (e.Button == MouseButtons.Left && crtButton.BackColor == DefaultBackColor)
            {
                Point mousePos = pnVideoEditing.PointToClient(Cursor.Position);
                int deltaX = (mousePos.X - mouseDownButtonX) - crtButton.Location.X;
                int deltaY = (mousePos.Y - mouseDownButtonY) - crtButton.Location.Y;

                if (deltaX == 0 && deltaY == 0) return;

                // Collect all items to move (selected + linked if enabled)
                HashSet<int> itemsToMove = new HashSet<int>(selectedItemIds);
                if (isLinkingEnabled)
                {
                    foreach (int id in selectedItemIds.ToList())
                    {
                        if (items.ContainsKey(id) && items[id].linkedItemId != -1)
                            itemsToMove.Add(items[id].linkedItemId);
                    }
                }

                // Apply snapping only to the primary dragged item for better UX
                int snappedX = crtButton.Location.X + deltaX;
                if (isSnappingEnabled)
                {
                    int bestSnap = snappedX;
                    int minDiff = SNAP_DISTANCE;

                    // Snap to other clips
                    foreach (var other in items.Values)
                    {
                        if (itemsToMove.Contains(GetCurrentKeyByButton(other.button))) continue;

                        int otherStart = (int)(other.startPoint / onePixelSec);
                        int otherEnd = (int)(other.endPoint / onePixelSec);

                        // Snap dragging start to other start/end
                        if (Math.Abs(snappedX - otherStart) < minDiff) { bestSnap = otherStart; minDiff = Math.Abs(snappedX - otherStart); }
                        if (Math.Abs(snappedX - otherEnd) < minDiff) { bestSnap = otherEnd; minDiff = Math.Abs(snappedX - otherEnd); }
                        
                        // Snap dragging end to other start/end
                        int draggedEnd = snappedX + crtButton.Width;
                        if (Math.Abs(draggedEnd - otherStart) < minDiff) { bestSnap = otherStart - crtButton.Width; minDiff = Math.Abs(draggedEnd - otherStart); }
                        if (Math.Abs(draggedEnd - otherEnd) < minDiff) { bestSnap = otherEnd - crtButton.Width; minDiff = Math.Abs(draggedEnd - otherEnd); }
                    }

                    // Snap to playhead
                    if (Math.Abs(snappedX - pbCursor.Location.X) < minDiff) { bestSnap = pbCursor.Location.X; minDiff = Math.Abs(snappedX - pbCursor.Location.X); }
                    if (Math.Abs((snappedX + crtButton.Width) - pbCursor.Location.X) < minDiff) { bestSnap = pbCursor.Location.X - crtButton.Width; minDiff = Math.Abs((snappedX + crtButton.Width) - pbCursor.Location.X); }

                    deltaX = bestSnap - crtButton.Location.X;
                }

                // Prevent moving before 0
                int minLocationX = int.MaxValue;
                foreach (int id in itemsToMove)
                {
                    if (items.ContainsKey(id))
                    {
                        if (items[id].button.Location.X < minLocationX)
                            minLocationX = items[id].button.Location.X;
                    }
                }
                if (minLocationX + deltaX < 0)
                {
                    deltaX = -minLocationX;
                }

                // Move all collected items
                foreach (int id in itemsToMove)
                {
                    if (items.ContainsKey(id))
                    {
                        Button btn = items[id].button;
                        btn.Location = new Point(btn.Location.X + deltaX, btn.Location.Y + (id == currentItem ? deltaY : 0));
                    }
                }
            }
        }

        private void Item_Button_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Button crtButton = sender as Button;
                int key = GetCurrentKeyByButton(crtButton);

                if (ModifierKeys == Keys.Control)
                {
                    if (selectedItemIds.Contains(key))
                        selectedItemIds.Remove(key);
                    else
                        selectedItemIds.Add(key);
                }
                else
                {
                    if (!selectedItemIds.Contains(key))
                    {
                        selectedItemIds.Clear();
                        selectedItemIds.Add(key);
                    }
                }

                Point position = pnVideoEditing.PointToClient(Cursor.Position);
                mouseDownButtonX = position.X - crtButton.Location.X;
                mouseDownButtonY = position.Y - crtButton.Location.Y;
                
                // Move cursor to click position if needed
                int logicalX = crtButton.Location.X + e.X;
                pbCursor.Location = new Point(logicalX, pbCursor.Location.Y);
                pbCursor.BringToFront();
                ChangeCursorLocation();

                currentItem = key;
                RefreshSelectionHighlight();
                
                // Update volume slider to match item
                // Update volume slider to match item
                if (items.ContainsKey(key))
                {
                     tbItemVolume.Value = (int)(items[key].Volume * 100);
                }
            }
            if (e.Button == MouseButtons.Right)
            {
                ShowItemContextMenu();
            }

        }

        private void RefreshSelectionHighlight()
        {
            foreach (var item in items.Values)
            {
                if (selectedItemIds.Contains(GetCurrentKeyByButton(item.button)))
                {
                    item.button.FlatAppearance.BorderSize = 2;
                    item.button.FlatAppearance.BorderColor = Color.Yellow;
                }
                else
                {
                    item.button.FlatAppearance.BorderSize = 0;
                    item.button.FlatAppearance.BorderColor = item.button.BackColor;
                }
            }
        }
        private void Item_Button_MouseUp(object sender, MouseEventArgs e)
        {
            Button crtButton = sender as Button;
            if (e.Button == MouseButtons.Left && crtButton.BackColor == DefaultBackColor)
            {
                // Commit changes for all moved items
                HashSet<int> itemsToMove = new HashSet<int>(selectedItemIds);
                if (isLinkingEnabled)
                {
                    foreach (int id in selectedItemIds.ToList())
                    {
                        if (items.ContainsKey(id) && items[id].linkedItemId != -1)
                            itemsToMove.Add(items[id].linkedItemId);
                    }
                }

                foreach (int key in itemsToMove)
                {
                    if (!items.ContainsKey(key)) continue;
                    var item = items[key];
                    var btn = item.button;

                    int interval = item.grid;
                    int centerY = btn.Location.Y + btn.Height / 2;
                    for (int i = 0; i < heightIntervals.Count() - 1; i++)
                    {
                        if (centerY >= heightIntervals[i] + INTERVAL_HEIGHT && centerY < heightIntervals[i + 1] + INTERVAL_HEIGHT)
                        {
                            interval = i;
                            break;
                        }
                    }
                    item.grid = interval;
                    item.startPoint = (double)btn.Location.X * onePixelSec;
                    item.endPoint = item.startPoint + item.duration;
                    btn.Location = new Point(btn.Location.X, interval * HEIGHT_LINES + INTERVAL_HEIGHT);
                }

                pnVideoEditing.Refresh();
            }
        }

        private void pnVideoEditing_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void pnVideoEditing_DragLeave(object sender, EventArgs e)
        {
        }

        private void Item_Button_Click(object sender, EventArgs e)
        {
            Button crt = sender as Button;
            int key = int.Parse(crt.Name.Split(',').Last());
            if (currentItem != key)
            {
                currentItem = key;
            }
        }
        public double Duration(string file)
        {
            IWMPMedia mediainfo = axWindowsMediaPlayer1.newMedia(file);
            return mediainfo.duration;
        }
        public string DurationReadeble(double duration)
        {
            return TimeSpan.FromSeconds(duration).ToString(@"mm\:ss");
        }
        private bool isPlaying()
        {
            return axWindowsMediaPlayer1.playState == WMPPlayState.wmppsPlaying;
        }
        private bool isPaused()
        {
            return axWindowsMediaPlayer1.playState == WMPPlayState.wmppsPaused;
        }
        private void FormEditor_ClientSizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                btnMaximize.Text = @"🗗";
            }
            else if (this.WindowState == FormWindowState.Normal)
            {
                btnMaximize.Text = @"🗖";
            }
            pnLeftMenu.Invalidate();
            pnVideoEditing.Invalidate();
        }
        private void pnVideoEditing_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Y < INTERVAL_HEIGHT)
            {
                Cursor = Cursors.HSplit;
                if (e.Button == MouseButtons.Left)
                {
                    // Calculate logical X by adding scroll value to client mouse X
                    int logicalX = e.X + pnVideoEditing.HorizontalScroll.Value; 
                    pbCursor.Location = new Point(logicalX, pbCursor.Location.Y);
                    pbCursor.BringToFront();
                    ChangeCursorLocation();
                    pnVideoEditing.Refresh();
                }
            }
            else
            {
                if (Cursor != Cursors.Default)
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void pnVideoEditing_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                // Calculate logical X by adding scroll value to client mouse X
                int logicalX = e.X + pnVideoEditing.HorizontalScroll.Value;
                pbCursor.Location = new Point(logicalX, pbCursor.Location.Y);
                pbCursor.BringToFront();
                ChangeCursorLocation();
                
                if (projectIsPlaying) // Pause on click/drag for better control
                {
                    projectIsPlaying = false;
                    axWindowsMediaPlayer1.Ctlcontrols.pause();
                    StopAudioPlayback();
                }

                if (isPreviewMode)
                {
                    isPreviewMode = false;
                    axWindowsMediaPlayer1.uiMode = "none";
                    axWindowsMediaPlayer1.Ctlcontrols.stop();
                }

                if (Cursor.Current == Cursors.Default)
                {
                    if (items.Count() > 0 && currentItem > 0)
                    {
                        UpdateCurrentItemByCursor();
                    }
                }
                pnVideoEditing.Refresh();
            }
        }
        private void FormEditor_MouseDown(object sender, MouseEventArgs e)
        {

        }
        private void FormEditor_MouseMove(object sender, MouseEventArgs e)
        {
        }
        private void FormEditor_KeyDown(object sender, KeyEventArgs e)
        {
            Debug.WriteLine("KeyDown: " + e.KeyCode);
            if (e.KeyCode == Keys.Space)
            {
                if (isPreviewMode)
                {
                    if (axWindowsMediaPlayer1.playState == WMPPlayState.wmppsPlaying)
                        axWindowsMediaPlayer1.Ctlcontrols.pause();
                    else
                        axWindowsMediaPlayer1.Ctlcontrols.play();
                    return;
                }
                projectIsPlaying = !projectIsPlaying;
                if (!projectIsPlaying) 
                {
                    axWindowsMediaPlayer1.Ctlcontrols.pause();
                    StopAudioPlayback();
                }
                else ChangeCursorLocation();
            }
            if (e.KeyCode == Keys.Delete)
            {
                DeleteCurrentItem();
            }
            if (e.Control && e.KeyCode == Keys.B)
            {
                SplitCurrentItem();
            }
        }
        void SplitCurrentItem()
        {
            double cursorSeconds = (double)pbCursor.Location.X * onePixelSec;
            var entry = items.FirstOrDefault(x => x.Value.startPoint < cursorSeconds && x.Value.endPoint > cursorSeconds);
            if (entry.Value != null)
            {
                int key = entry.Key;
                Item item = entry.Value;
                double oldTotalDuration = item.duration;
                double relativeSplitTime = cursorSeconds - item.startPoint;
                double splitRatio = relativeSplitTime / oldTotalDuration;

                int newKey = items.Keys.Max() + 1;
                Item secondHalf = new Item
                {
                    path = item.path,
                    fileName = item.fileName,
                    type = item.type,
                    fileType = item.fileType,
                    grid = item.grid,
                    duration = oldTotalDuration - relativeSplitTime,
                    startPoint = cursorSeconds,
                    startVideo = item.startVideo + relativeSplitTime,
                    endVideo = item.endVideo
                };
                secondHalf.endPoint = secondHalf.startPoint + secondHalf.duration;

                Bitmap bmp1 = null, bmp2 = null;
                // Crop Background Image for visual accuracy
                if (item.button.BackgroundImage != null)
                {
                    Image fullImg = item.button.BackgroundImage;
                    int splitPixel = (int)Math.Round((double)item.button.Width * splitRatio);
                    splitPixel = Math.Max(1, Math.Min(item.button.Width - 1, splitPixel));

                    // Calculate how much of the source image each pixel represents
                    double pixelToImageRatio = (double)fullImg.Width / (double)item.button.Width;
                    int srcSplitX = (int)Math.Round(splitPixel * pixelToImageRatio);
                    srcSplitX = Math.Max(1, Math.Min(fullImg.Width - 1, srcSplitX));

                    bmp1 = new Bitmap(srcSplitX, fullImg.Height);
                    using (Graphics g = Graphics.FromImage(bmp1))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(fullImg, new Rectangle(0, 0, srcSplitX, fullImg.Height), 
                            new Rectangle(0, 0, srcSplitX, fullImg.Height), GraphicsUnit.Pixel);
                    }

                    bmp2 = new Bitmap(fullImg.Width - srcSplitX, fullImg.Height);
                    using (Graphics g = Graphics.FromImage(bmp2))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(fullImg, new Rectangle(0, 0, fullImg.Width - srcSplitX, fullImg.Height), 
                            new Rectangle(srcSplitX, 0, fullImg.Width - srcSplitX, fullImg.Height), GraphicsUnit.Pixel);
                    }

                    item.button.BackgroundImage = bmp1;
                    item.button.BackgroundImageLayout = ImageLayout.Stretch;
                }

                // Update first half
                item.duration = relativeSplitTime;
                item.endPoint = cursorSeconds;
                item.endVideo = item.startVideo + relativeSplitTime;
                item.button.Text = item.fileName + " : " + DurationReadeble(item.duration);

                // Create new button for second half
                Button btnSecond = new Button();
                btnSecond.FlatStyle = FlatStyle.Flat;
                btnSecond.FlatAppearance.BorderSize = 0;
                btnSecond.BackColor = item.button.BackColor;
                btnSecond.Height = HEIGHT_LINES;
                btnSecond.Text = secondHalf.fileName + " : " + DurationReadeble(secondHalf.duration);
                btnSecond.Name = secondHalf.path + "," + secondHalf.duration + "," + newKey;
                btnSecond.Cursor = Cursors.Hand;
                btnSecond.Click += Item_Button_Click;
                btnSecond.MouseUp += Item_Button_MouseUp;
                btnSecond.MouseDown += Item_Button_MouseDown;
                btnSecond.MouseMove += Item_Button_MouseMove;
                
                if (bmp2 != null)
                {
                    btnSecond.BackgroundImage = bmp2;
                    btnSecond.BackgroundImageLayout = ImageLayout.Stretch;
                }

                pnVideoEditing.Controls.Add(btnSecond);
                secondHalf.button = btnSecond;
                secondHalf.Volume = item.Volume; // Copy volume settings
                items.Add(newKey, secondHalf);

                RefreshItemsPositionAndSize();
                currentItem = newKey;
                LoadItem(newKey);
                axWindowsMediaPlayer1.Ctlcontrols.currentPosition = secondHalf.startVideo;
                if (projectIsPlaying) axWindowsMediaPlayer1.Ctlcontrols.play();
                else axWindowsMediaPlayer1.Ctlcontrols.pause();
                pnVideoEditing.Update();
            }
        }
        void DeleteCurrentItem()
        {
            if (selectedItemIds.Count > 0)
            {
                if (isPlaying())
                {
                    axWindowsMediaPlayer1.Ctlcontrols.pause();
                }

                foreach (int id in selectedItemIds.ToList())
                {
                    if (items.ContainsKey(id))
                    {
                        Button btn = items[id].button;
                        pnVideoEditing.Controls.Remove(btn);
                        items.Remove(id);
                    }
                }

                selectedItemIds.Clear();
                if (items.Count > 0)
                {
                    currentItem = items.Keys.Max();
                }
                else
                {
                    currentItem = 0;
                }
                pnVideoEditing.Refresh();
            }
        }
        private void FormEditor_KeyPress(object sender, KeyPressEventArgs e)
        {
            Debug.WriteLine("KeyPress: " + e.KeyChar);
        }

        private void FormEditor_SizeChanged(object sender, EventArgs e)
        {
            pnLeftMenu.Invalidate();
            pnVideoEditing.Invalidate();
        }

        private void axWindowsMediaPlayer1_Enter(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!projectIsPlaying && !isPreviewMode) return;
            if (isPreviewMode) return;

            double currentTime = (double)pbCursor.Location.X * onePixelSec;
            // Use strict inequality for endPoint to avoid edge case where cursor is exactly at end
            
            // Should prioritize items based on track type (Video > Audio) for playback control of WMP
            // But we need to check bounds.
            var intersectingItems = items.Values.Where(x => x.startPoint <= currentTime && x.endPoint > currentTime).ToList();
            
            // Find video item if any, otherwise just take first one (likely audio)
            var crt = intersectingItems.FirstOrDefault(x => x.fileType == FileType.Video) ?? intersectingItems.FirstOrDefault();

            if (crt != null)
            {
                int key = items.FirstOrDefault(x => x.Value == crt).Key;
                
                // Check if we're about to reach the end of this clip
                double timeInClip = currentTime - crt.startPoint;
                double videoPosition = crt.startVideo + timeInClip;
                
                if (videoPosition >= crt.endVideo - 0.05)
                {
                    // At the end of the clip - pause and move cursor past the end point
                    axWindowsMediaPlayer1.Ctlcontrols.pause();
                    // Position the cursor just beyond this clip's end to trigger gap mode
                    double timeDelta = (double)cursorTimer.Interval / 1000.0;
                    pbCursor.Location = new Point((int)Math.Round((crt.endPoint + timeDelta) / onePixelSec), pbCursor.Location.Y);
                }
                else
                {
                    // Normal playback within the clip
                    if (currentItem != key || axWindowsMediaPlayer1.playState != WMPPlayState.wmppsPlaying)
                    {
                        if (axWindowsMediaPlayer1.URL != crt.path) LoadItem(key);
                        currentItem = key;
                        axWindowsMediaPlayer1.Ctlcontrols.currentPosition = videoPosition;
                        axWindowsMediaPlayer1.Ctlcontrols.play();
                        
                        if (pnBlackScreen != null) pnBlackScreen.Visible = false;
                    }
                     
                     // Keep cursor synced with actual video position
                    // Only sync if video is actually playing to avoid jumps during buffering
                    if (axWindowsMediaPlayer1.playState == WMPPlayState.wmppsPlaying)
                    {
                         double actualGlobalTime = crt.startPoint + (axWindowsMediaPlayer1.Ctlcontrols.currentPosition - crt.startVideo);
                         pbCursor.Location = new Point((int)Math.Round(actualGlobalTime / onePixelSec), pbCursor.Location.Y);
                    }
                    else
                    {
                         // If buffering or staring, rely on timer estimation
                         // Or if audio only, we just move cursor manually
                         if (crt.fileType != FileType.Video)
                         {
                              double timeDelta = (double)cursorTimer.Interval / 1000.0;
                              pbCursor.Location = new Point((int)Math.Round((currentTime + timeDelta) / onePixelSec), pbCursor.Location.Y);
                         }
                    }
                }
            }
            else
            {
                // In a gap between clips
                if (axWindowsMediaPlayer1.playState == WMPPlayState.wmppsPlaying)
                {
                    axWindowsMediaPlayer1.Ctlcontrols.pause();
                }

                if (pnBlackScreen != null)
                {
                    pnBlackScreen.BringToFront();
                    pnBlackScreen.Visible = true;
                }

                double timeDelta = (double)cursorTimer.Interval / 1000.0;
                double nextTime = currentTime + timeDelta;
                pbCursor.Location = new Point((int)Math.Round(nextTime / onePixelSec), pbCursor.Location.Y);
            }

            // Global Update for Audio and Time
            double displayTime = (double)pbCursor.Location.X * onePixelSec;
            if (lblCurrentTime != null)
                lblCurrentTime.Text = TimeSpan.FromSeconds(displayTime).ToString(@"hh\:mm\:ss\:ff");
                
            UpdateAudioPlayback(displayTime);
        }

        private void ShowItemContextMenu()
        {
            if (items.Count > 0)
            {
                ContextMenu cm = new ContextMenu();
                if (isPaused())
                {
                    cm.MenuItems.Add("Split", (s, ev) => SplitCurrentItem());
                }
                
                // Get item from selection logic or cursor logic
                Item crtItem = null;
                if (selectedItemIds.Count > 0) 
                     crtItem = items[selectedItemIds[0]];
                else 
                     crtItem = GetIntersectedItem();

                if (crtItem != null && crtItem.fileType == FileType.Video)
                {
                    cm.MenuItems.Add("Detach Audio", new EventHandler(DetachAudio));
                }
                if (crtItem != null && crtItem.fileType == FileType.Audio)
                {
                    MenuItem volItem = new MenuItem("Volume (Current: " + (int)(crtItem.Volume * 100) + "%)");
                    volItem.MenuItems.Add("Mute", (s, ev) => { UpdateItemVolume(crtItem, 0f); tbItemVolume.Value = 0; });
                    volItem.MenuItems.Add("25%", (s, ev) => { UpdateItemVolume(crtItem, 0.25f); tbItemVolume.Value = 25; });
                    volItem.MenuItems.Add("50%", (s, ev) => { UpdateItemVolume(crtItem, 0.50f); tbItemVolume.Value = 50; });
                    volItem.MenuItems.Add("75%", (s, ev) => { UpdateItemVolume(crtItem, 0.75f); tbItemVolume.Value = 75; });
                    volItem.MenuItems.Add("100%", (s, ev) => { UpdateItemVolume(crtItem, 1.0f); tbItemVolume.Value = 100; });
                    volItem.MenuItems.Add("125%", (s, ev) => { UpdateItemVolume(crtItem, 1.25f); tbItemVolume.Value = 125; });
                    
                    volItem.MenuItems.Add("-");
                    volItem.MenuItems.Add("Custom Value...", (s, ev) => {
                         ShowCustomVolumeDialog(crtItem);
                    });
                    
                    cm.MenuItems.Add(volItem);
                }
                cm.MenuItems.Add("Remove", new EventHandler(RemoveFile));
                cm.Show(pnVideoEditing, pnVideoEditing.PointToClient(Cursor.Position));
            }
        }

        private void ShowCustomVolumeDialog(Item item)
        {
            Form prompt = new Form()
            {
                Width = 250,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Set Volume %",
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };
            
            Label textLabel = new Label() { Left = 20, Top = 20, Text = "Enter Volume % (0-200):", AutoSize = true };
            TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 190, Text = ((int)(item.Volume * 100)).ToString() };
            Button confirmation = new Button() { Text = "Ok", Left = 130, Width = 80, Top = 80, DialogResult = DialogResult.OK };
            
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                if (int.TryParse(textBox.Text, out int result))
                {
                    if (result < 0) result = 0;
                    if (result > 200) result = 200; // Cap at 200%
                    float vol = result / 100f;
                    UpdateItemVolume(item, vol);
                    if (tbItemVolume.Maximum < result) tbItemVolume.Maximum = result; 
                    if (result <= 125) tbItemVolume.Maximum = 125;
                    tbItemVolume.Value = Math.Min(tbItemVolume.Maximum, result);
                }
            }
        }
        /// <summary>
        /// MAIN CURSOR
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pbCursor_Click(object sender, EventArgs e)
        {

        }

        private void pbCursor_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                // Convert screen mouse position to logical timeline position
                Point clientPos = pnVideoEditing.PointToClient(Cursor.Position);
                int logicalX = clientPos.X + pnVideoEditing.HorizontalScroll.Value;
                
                pbCursor.Location = new Point(logicalX, pbCursor.Location.Y);
                pbCursor.BringToFront();
                pbCursor.Update();
            }
        }
        int GetCurrentKeyByButton(Button btn)
        {
            return int.Parse(btn.Name.Split(',').Last());
        }
        Item GetIntersectedItem()
        {
            Button crtButton = GetIntersectedButton();
            if (crtButton == null)
            {
                return null;
            }
            int key = GetCurrentKeyByButton(crtButton);
            return items[key];
        }
        Button GetIntersectedButton()
        {
            foreach (var btn in pnVideoEditing.Controls.OfType<Button>())
            {
                if (pbCursor.Bounds.IntersectsWith(btn.Bounds))
                {
                    return btn;
                }
            }
            return null;
        }
        void UpdateCurrentItemByCursor()
        {
            double cursorSeconds = pbCursor.Location.X * onePixelSec;
            var intersect = items.Values.Where(x => x.startPoint <= cursorSeconds && x.endPoint >= cursorSeconds).ToList();
            if (intersect.Count == 0) return;
            
            // Prioritize Video for WMP playback
            var crt = intersect.FirstOrDefault(x => x.fileType == FileType.Video) ?? intersect.First();
            
            int key = items.FirstOrDefault(x => x.Value == crt).Key;
            currentItem = key;
            LoadItem(currentItem);
            axWindowsMediaPlayer1.Ctlcontrols.currentPosition = cursorSeconds - items[currentItem].startPoint;
            axWindowsMediaPlayer1.Ctlcontrols.play();
        }
        void ChangeCursorLocation()
        {
            double cursorSeconds = (double)pbCursor.Location.X * onePixelSec;
            var intersect = items.Values.Where(x => x.startPoint <= cursorSeconds && x.endPoint >= cursorSeconds).ToList();

            if (intersect.Count > 0)
            {
                // Prioritize Video
                var item = intersect.FirstOrDefault(x => x.fileType == FileType.Video) ?? intersect.First();
                int key = items.FirstOrDefault(x => x.Value == item).Key;

                if (currentItem != key || axWindowsMediaPlayer1.URL != item.path)
                {
                    LoadItem(key);
                    currentItem = key;
                }
                axWindowsMediaPlayer1.Ctlcontrols.currentPosition = item.startVideo + (cursorSeconds - item.startPoint);
                if (projectIsPlaying) axWindowsMediaPlayer1.Ctlcontrols.play();
                else 
                {
                    axWindowsMediaPlayer1.Ctlcontrols.pause();
                    StopAudioPlayback();
                }
                
                if (pnBlackScreen != null) pnBlackScreen.Visible = false;
            }
            else
            {
                // In a gap
                axWindowsMediaPlayer1.Ctlcontrols.pause();
                StopAudioPlayback();
                currentItem = 0;
                if (pnBlackScreen != null) 
                {
                    pnBlackScreen.BringToFront();
                    pnBlackScreen.Visible = true;
                }
            }
        }
        private void pbCursor_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ChangeCursorLocation();
            }
            if (e.Button == MouseButtons.Right)
            {
                // Check if we are in the timeline header
                if (e.Location.Y < INTERVAL_HEIGHT)
                {
                    ContextMenu cm = new ContextMenu();
                    cm.MenuItems.Add("Set Render Start", (s, ev) => {
                        renderStartTime = (double)pbCursor.Location.X * onePixelSec;
                        pnVideoEditing.Invalidate();
                    });
                    cm.MenuItems.Add("Set Render End", (s, ev) => {
                        renderEndTime = (double)pbCursor.Location.X * onePixelSec;
                        pnVideoEditing.Invalidate();
                    });
                    cm.MenuItems.Add("Clear Render Section", (s, ev) => {
                        renderStartTime = -1;
                        renderEndTime = -1;
                        pnVideoEditing.Invalidate();
                    });
                    cm.Show(pnVideoEditing, e.Location);
                }
                else
                {
                    ShowItemContextMenu();
                }
            }
        }

        private void UpdateItemVolume(Item item, float volume)
        {
            item.Volume = volume;
            // Immediate update if playing
            int key = items.FirstOrDefault(x => x.Value == item).Key;
            if (activeAudioStreams.ContainsKey(key))
            {
                activeAudioStreams[key].Volume = volume;
            }
        }

        private void DetachAudio(object sender, EventArgs e)
        {
            Item intItem = GetIntersectedItem();
            if (intItem == null) return;
            int videoKey = GetCurrentKeyByButton(intItem.button);
            var wavFileName = VideoProcessing.FileNameWithoutExtension(intItem.fileName) + ".wav";
            var output = currentProject.projectsPath + currentProject.projectName + @"\" + wavFileName;
            VideoProcessing.ConvertFileToWAV(intItem.path, (int)intItem.duration, output);
            ExtractVideoWaveFile(output, wavFileName, intItem.duration,
                VideoProcessing.CreateWaveImage(output, (int)intItem.duration), ".wav", intItem.startPoint, intItem.grid + 1, videoKey);
        }

        private void ExtractVideoWaveFile(string file, string fileName, double duration, Image waveImage, string extension, double startPos, int nextGrid, int sourceKey = -1)
        {
            if (extension == ".wav")
            {
                int key = 0;
                if (items.Count > 0)
                {
                    key = items.Keys.Max() + 1;
                }
                else
                {
                    key = 1;
                }

                Button crtButton = new Button();
                crtButton.FlatStyle = FlatStyle.Flat;
                crtButton.FlatAppearance.BorderSize = 0;
                crtButton.Text = fileName + " : " + DurationReadeble(duration);
                crtButton.Name = file + "," + duration + "," + key;
                crtButton.Cursor = Cursors.Hand;
                crtButton.Click += Item_Button_Click;
                crtButton.MouseUp += Item_Button_MouseUp;
                crtButton.MouseDown += Item_Button_MouseDown;
                crtButton.MouseMove += Item_Button_MouseMove;
                crtButton.Size = new Size((int)Math.Round(duration / onePixelSec), HEIGHT_LINES);
                crtButton.Height = HEIGHT_LINES;
                crtButton.BackgroundImageLayout = ImageLayout.Stretch;
                crtButton.BackgroundImage = waveImage;
                pnVideoEditing.Controls.Add(crtButton);
                crtButton.Location = new Point((int)Math.Round(startPos / onePixelSec), nextGrid * HEIGHT_LINES + INTERVAL_HEIGHT);
                //pbCursor.Location = new Point(crtButton.Location.X, pbCursor.Location.Y);
                Item item = new Item
                {
                    path = file,
                    fileName = fileName,
                    type = extension,
                    fileType = FileType.Audio,
                    grid = nextGrid,
                    startPoint = startPos,
                    endPoint = startPos + duration,
                    startVideo = 0,
                    endVideo = duration,
                    duration = duration,
                    button = crtButton,
                };
                items.Add(key, item);
                
                if (sourceKey != -1 && items.ContainsKey(sourceKey))
                {
                    items[sourceKey].linkedItemId = key;
                    items[key].linkedItemId = sourceKey;
                }

                LoadItem(key);
                currentItem = key;
                pnVideoEditing.Update();
            }
        }

        private void RemoveFile(object sender, EventArgs e)
        {
            DeleteCurrentItem();
        }


        void LoadItem(int key)
        {
            if (items.Count > 0 && key > 0)
            {
                var item = items[key];
                if (axWindowsMediaPlayer1.URL != item.path)
                {
                    axWindowsMediaPlayer1.URL = item.path;
                }
                
                // Logic to prevent double audio:
                // If the item is AUDIO, WMP should be muted (NAudio handles it).
                // If the item is VIDEO, we now ALWAYS mute WMP. This is because we architecture requires
                // separate audio tracks for mixing. If the user deletes that track, they want silence.
                
                bool handledByNAudio = false;
                if (item.fileType == FileType.Audio) 
                {
                    handledByNAudio = true;
                }
                else if (item.fileType == FileType.Video)
                {
                    // Always mute WMP for video files to rely 100% on the extracted audio track
                    handledByNAudio = true;
                }

                if (handledByNAudio)
                {
                    axWindowsMediaPlayer1.settings.mute = true;
                }
                else
                {
                    axWindowsMediaPlayer1.settings.mute = false;
                    axWindowsMediaPlayer1.settings.volume = tbVolume.Value;
                }
            }
        }

        private void FormEditor_Resize(object sender, EventArgs e)
        {
            visibleSize = new Size(this.Width - SystemInformation.FixedFrameBorderSize.Width * 2 - pnVideoEditing.Location.X - 15, this.Height - (splitHorizontal.Location.Y + splitHorizontal.SplitterDistance + this.DefaultMargin.Vertical));
            UpdateTotalViewNrOfPixels();
            pnVideoEditing.Refresh();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void pnTitle_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, 0x112, 0xf012, 0);
            if (e.Button == MouseButtons.Left && e.Clicks >= 2)
            {
                pnTitle_MouseDoubleClick(sender, e);
                return;
            }

            if (e.Button != MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, Constants.WM_NCLBUTTONDOWN, Constants.HTCAPTION, 0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {

                this.WindowState = FormWindowState.Maximized;
                btnMaximize.Text = @"🗗";
            }
            else if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
                btnMaximize.Text = @"🗖";
            }
            if (items.Count > 0)
            {
                items[currentItem].button.Select();
            }
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Minimized;
            }
            if (items.Count > 0)
            {
                items[currentItem].button.Select();
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Project test = new Project(false);
            if (test.projectName != "")
            {
                currentProject = test;
                lblProjectName.Text = "Current Project: " + currentProject.projectName;
                lblProjectName.Location = new Point(this.Width / 2 - lblProjectName.Width / 2, lblProjectName.Location.Y);
            }
        }

        private void btnExit_MouseEnter(object sender, EventArgs e)
        {
            Button crtButton = sender as Button;
            crtButton.BackColor = Color.Red;
        }

        private void btnExit_MouseLeave(object sender, EventArgs e)
        {
            Button crtButton = sender as Button;
            crtButton.BackColor = Color.Transparent;
        }

        private void pnTitle_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.WindowState == FormWindowState.Normal)
                {
                    this.WindowState = FormWindowState.Maximized;
                    btnMaximize.Text = @"🗗";
                }
                else if (this.WindowState == FormWindowState.Maximized)
                {
                    this.WindowState = FormWindowState.Normal;
                    btnMaximize.Text = @"🗖";
                }
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "Video Editor v" + Application.ProductVersion +
                "\nCreated by Tutorialeu.ro\nAll rights Reserved!");
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "Press Space to pause!\nRight click on cursor for other options!\n" +
                "Split is avaliable only on pause time!\nPress delete to remove a video from editor!\nSelect and move the videos as you want!");
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteCurrentItem();
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = currentProject.projectsPath;
                openFileDialog.Filter = "json files (*.json)|*.json";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = string.Empty;
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;


                    currentProject = new Project(filePath);
                    lblProjectName.Text = "Current Project: " + currentProject.projectName;
                    lblProjectName.Location = new Point(this.Width / 2 - lblProjectName.Width / 2, lblProjectName.Location.Y);
                }
            }
        }

        private void hScrollBarZoom_ValueChanged(object sender, EventArgs e)
        {
            Console.WriteLine("Scroll Value Changed: " + hScrollBarZoom.Value);
            if (stage != hScrollBarZoom.Value)
            {
                fragment = MIN_FRAGMENT;
                stage = hScrollBarZoom.Value;
                onePixelSec = stageIntervals[stage] / (fragment * 20d);
                RefreshItemsPositionAndSize();
                pnVideoEditing.Refresh();
            }
        }

        private void lblProjectName_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, 0x112, 0xf012, 0);
            if (e.Button == MouseButtons.Left && e.Clicks >= 2)
            {
                pnTitle_MouseDoubleClick(sender, e);
                return;
            }

            if (e.Button != MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, Constants.WM_NCLBUTTONDOWN, Constants.HTCAPTION, 0);
            }
        }

        private void pnVideoEditing_Resize(object sender, EventArgs e)
        {
        }

        private void pnVideoEditing_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }
    }
}
