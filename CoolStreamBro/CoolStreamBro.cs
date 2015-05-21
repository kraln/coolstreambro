/*
 * Cool Stream, Bro
 * 
 * This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
 * If a copy of the MPL was not distributed with this file, 
 * You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DirectXCaptureWrapper.Capture;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace CoolStreamBro
{
  public class CoolStreamBro : System.Windows.Forms.Form
  {
    private Capture capture = null;
    private Filters filters = new Filters();

    private TextBox tFilename;
    private Label lFilenameLabel;
    private Button bStart;
    private Button bStop;
    private MenuItem mOptions;
    private MainMenu mainMenu;
    private MenuItem mInputs;
    private MenuItem mVideoDevices;
    private MenuItem mAudioDevices;
    private MenuItem mVideoCompressors;
    private MenuItem mAudioCompressors;
    private MenuItem mVideoSources;
    private MenuItem mAudioSources;
    private Panel pPreviewVideo;
    private MenuItem menuItem4;
    private MenuItem mAudioChannels;
    private MenuItem mAudioSamplingRate;
    private MenuItem mAudioSampleSizes;
    private MenuItem menuItem5;
    private MenuItem mFrameSizes;
    private MenuItem mFrameRates;
    private MenuItem menuItem6;
    private MenuItem mPreviewEnable;
    private MenuItem mPropertyPages;
    private MenuItem mVideoCaps;
    private MenuItem mAudioCaps;
    private MenuItem mChannel;
    private MenuItem menuItem3;
    private MenuItem mInputType;
    private MenuItem mPreview;
    private MenuItem menuItem9;
    private MenuItem mPreviewAudio;
    private MenuItem mPreviewAspect;
    private IContainer components;
    private float aspect;
    private MenuItem mFreeAspect;
    private MenuItem mSquareAspect;
    private MenuItem m43Aspect;
    private MenuItem m169Aspect;
    private bool fixed_aspect = false;

    public CoolStreamBro()
    {

      InitializeComponent();
      try
      {
        updateMenu();
      }
      catch
      {
      }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
      public int Left;
      public int Top;
      public int Right;
      public int Bottom;
    }

    /* Override the windows messages so that we can enforce some ratios */
    protected override void WndProc(ref Message m)
    {

      /* if fixed aspect ratio is desired */
      if (fixed_aspect && (m.Msg == 0x216 || m.Msg == 0x214))
      { // WM_MOVING || WM_SIZING

        RECT rc = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));

        int w = rc.Right - rc.Left;
        int h = rc.Bottom - rc.Top;
        int z = w > h ? w : h;
        rc.Bottom = rc.Top + z;
        rc.Right = rc.Left + (int)(z * aspect);
        Marshal.StructureToPtr(rc, m.LParam, false);
        m.Result = (IntPtr)1;
        return;
      }
      base.WndProc(ref m);
    }

    protected override void Dispose(bool disposing)
    {
      if (disposing)
      {
        if (components != null)
        {
          components.Dispose();
        }
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
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CoolStreamBro));
      this.tFilename = new System.Windows.Forms.TextBox();
      this.lFilenameLabel = new System.Windows.Forms.Label();
      this.bStart = new System.Windows.Forms.Button();
      this.bStop = new System.Windows.Forms.Button();
      this.mainMenu = new System.Windows.Forms.MainMenu(this.components);
      this.mInputs = new System.Windows.Forms.MenuItem();
      this.mVideoDevices = new System.Windows.Forms.MenuItem();
      this.mAudioDevices = new System.Windows.Forms.MenuItem();
      this.menuItem4 = new System.Windows.Forms.MenuItem();
      this.mVideoCompressors = new System.Windows.Forms.MenuItem();
      this.mAudioCompressors = new System.Windows.Forms.MenuItem();
      this.mOptions = new System.Windows.Forms.MenuItem();
      this.mVideoSources = new System.Windows.Forms.MenuItem();
      this.mFrameSizes = new System.Windows.Forms.MenuItem();
      this.mFrameRates = new System.Windows.Forms.MenuItem();
      this.menuItem5 = new System.Windows.Forms.MenuItem();
      this.mAudioSources = new System.Windows.Forms.MenuItem();
      this.mAudioChannels = new System.Windows.Forms.MenuItem();
      this.mAudioSamplingRate = new System.Windows.Forms.MenuItem();
      this.mAudioSampleSizes = new System.Windows.Forms.MenuItem();
      this.menuItem3 = new System.Windows.Forms.MenuItem();
      this.mChannel = new System.Windows.Forms.MenuItem();
      this.mInputType = new System.Windows.Forms.MenuItem();
      this.menuItem6 = new System.Windows.Forms.MenuItem();
      this.mPropertyPages = new System.Windows.Forms.MenuItem();
      this.mVideoCaps = new System.Windows.Forms.MenuItem();
      this.mAudioCaps = new System.Windows.Forms.MenuItem();
      this.mPreview = new System.Windows.Forms.MenuItem();
      this.mPreviewEnable = new System.Windows.Forms.MenuItem();
      this.menuItem9 = new System.Windows.Forms.MenuItem();
      this.mPreviewAudio = new System.Windows.Forms.MenuItem();
      this.mPreviewAspect = new System.Windows.Forms.MenuItem();
      this.mFreeAspect = new System.Windows.Forms.MenuItem();
      this.mSquareAspect = new System.Windows.Forms.MenuItem();
      this.m43Aspect = new System.Windows.Forms.MenuItem();
      this.m169Aspect = new System.Windows.Forms.MenuItem();
      this.pPreviewVideo = new System.Windows.Forms.Panel();
      this.pPreviewVideo.SuspendLayout();
      this.SuspendLayout();
      // 
      // tFilename
      // 
      this.tFilename.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.tFilename.Location = new System.Drawing.Point(107, 371);
      this.tFilename.Name = "tFilename";
      this.tFilename.Size = new System.Drawing.Size(192, 20);
      this.tFilename.TabIndex = 0;
      this.tFilename.Text = "c:\\users\\";
      this.tFilename.TextChanged += new System.EventHandler(this.txtFilename_TextChanged);
      // 
      // lFilenameLabel
      // 
      this.lFilenameLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.lFilenameLabel.AutoSize = true;
      this.lFilenameLabel.Location = new System.Drawing.Point(5, 372);
      this.lFilenameLabel.Name = "lFilenameLabel";
      this.lFilenameLabel.Size = new System.Drawing.Size(92, 13);
      this.lFilenameLabel.TabIndex = 1;
      this.lFilenameLabel.Text = "Capture Filename:";
      this.lFilenameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
      // 
      // bStart
      // 
      this.bStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.bStart.Location = new System.Drawing.Point(305, 368);
      this.bStart.Name = "bStart";
      this.bStart.Size = new System.Drawing.Size(80, 24);
      this.bStart.TabIndex = 2;
      this.bStart.Text = "Start Capture";
      this.bStart.Click += new System.EventHandler(this.bStart_Click);
      // 
      // bStop
      // 
      this.bStop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.bStop.Location = new System.Drawing.Point(391, 368);
      this.bStop.Name = "bStop";
      this.bStop.Size = new System.Drawing.Size(80, 24);
      this.bStop.TabIndex = 3;
      this.bStop.Text = "Stop Capture";
      this.bStop.Click += new System.EventHandler(this.bStop_Click);
      // 
      // mainMenu
      // 
      this.mainMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mInputs,
            this.mOptions,
            this.mPreview});
      // 
      // mInputs
      // 
      this.mInputs.Index = 0;
      this.mInputs.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mVideoDevices,
            this.mAudioDevices,
            this.menuItem4,
            this.mVideoCompressors,
            this.mAudioCompressors});
      this.mInputs.Text = "Inputs";
      // 
      // mVideoDevices
      // 
      this.mVideoDevices.Index = 0;
      this.mVideoDevices.Text = "Video Devices";
      // 
      // mAudioDevices
      // 
      this.mAudioDevices.Index = 1;
      this.mAudioDevices.Text = "Audio Devices";
      // 
      // menuItem4
      // 
      this.menuItem4.Index = 2;
      this.menuItem4.Text = "-";
      // 
      // mVideoCompressors
      // 
      this.mVideoCompressors.Index = 3;
      this.mVideoCompressors.Text = "Video Compressors";
      // 
      // mAudioCompressors
      // 
      this.mAudioCompressors.Index = 4;
      this.mAudioCompressors.Text = "Audio Compressors";
      // 
      // mOptions
      // 
      this.mOptions.Index = 1;
      this.mOptions.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mVideoSources,
            this.mFrameSizes,
            this.mFrameRates,
            this.menuItem5,
            this.mAudioSources,
            this.mAudioChannels,
            this.mAudioSamplingRate,
            this.mAudioSampleSizes,
            this.menuItem3,
            this.mChannel,
            this.mInputType,
            this.menuItem6,
            this.mPropertyPages,
            this.mVideoCaps,
            this.mAudioCaps});
      this.mOptions.Text = "Options";
      // 
      // mVideoSources
      // 
      this.mVideoSources.Index = 0;
      this.mVideoSources.Text = "Video Sources";
      // 
      // mFrameSizes
      // 
      this.mFrameSizes.Index = 1;
      this.mFrameSizes.Text = "Resolution";
      // 
      // mFrameRates
      // 
      this.mFrameRates.Index = 2;
      this.mFrameRates.Text = "FPS";
      this.mFrameRates.Click += new System.EventHandler(this.mFrameRates_Click);
      // 
      // menuItem5
      // 
      this.menuItem5.Index = 3;
      this.menuItem5.Text = "-";
      // 
      // mAudioSources
      // 
      this.mAudioSources.Index = 4;
      this.mAudioSources.Text = "Audio Sources";
      // 
      // mAudioChannels
      // 
      this.mAudioChannels.Index = 5;
      this.mAudioChannels.Text = "Mix";
      // 
      // mAudioSamplingRate
      // 
      this.mAudioSamplingRate.Index = 6;
      this.mAudioSamplingRate.Text = "Sample Rate";
      // 
      // mAudioSampleSizes
      // 
      this.mAudioSampleSizes.Index = 7;
      this.mAudioSampleSizes.Text = "Sample Size";
      // 
      // menuItem3
      // 
      this.menuItem3.Index = 8;
      this.menuItem3.Text = "-";
      // 
      // mChannel
      // 
      this.mChannel.Index = 9;
      this.mChannel.Text = "Tuner Input";
      // 
      // mInputType
      // 
      this.mInputType.Index = 10;
      this.mInputType.Text = "Tuner Type";
      this.mInputType.Click += new System.EventHandler(this.mInputType_Click);
      // 
      // menuItem6
      // 
      this.menuItem6.Index = 11;
      this.menuItem6.Text = "-";
      // 
      // mPropertyPages
      // 
      this.mPropertyPages.Index = 12;
      this.mPropertyPages.Text = "Input Properties";
      // 
      // mVideoCaps
      // 
      this.mVideoCaps.Index = 13;
      this.mVideoCaps.Text = "Video Caps";
      this.mVideoCaps.Click += new System.EventHandler(this.mVideoCaps_Click);
      // 
      // mAudioCaps
      // 
      this.mAudioCaps.Index = 14;
      this.mAudioCaps.Text = "Audio Caps";
      this.mAudioCaps.Click += new System.EventHandler(this.mAudioCaps_Click);
      // 
      // mPreview
      // 
      this.mPreview.Index = 2;
      this.mPreview.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mPreviewEnable,
            this.menuItem9,
            this.mPreviewAudio,
            this.mPreviewAspect});
      this.mPreview.Text = "Preview";
      // 
      // mPreviewEnable
      // 
      this.mPreviewEnable.Index = 0;
      this.mPreviewEnable.Text = "Preview Enabled";
      this.mPreviewEnable.Click += new System.EventHandler(this.mPreview_Click);
      // 
      // menuItem9
      // 
      this.menuItem9.Index = 1;
      this.menuItem9.Text = "-";
      // 
      // mPreviewAudio
      // 
      this.mPreviewAudio.Index = 2;
      this.mPreviewAudio.Text = "Include Audio";
      this.mPreviewAudio.Click += new System.EventHandler(this.mPreviewAudio_Click);
      // 
      // mPreviewAspect
      // 
      this.mPreviewAspect.Index = 3;
      this.mPreviewAspect.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mFreeAspect,
            this.mSquareAspect,
            this.m43Aspect,
            this.m169Aspect});
      this.mPreviewAspect.Text = "Fixed Aspect";
      // 
      // mFreeAspect
      // 
      this.mFreeAspect.Checked = true;
      this.mFreeAspect.Index = 0;
      this.mFreeAspect.Text = "Not Fixed";
      this.mFreeAspect.Click += new System.EventHandler(this.mFreeAspect_Click);
      // 
      // mSquareAspect
      // 
      this.mSquareAspect.Index = 1;
      this.mSquareAspect.Text = "1:1";
      this.mSquareAspect.Click += new System.EventHandler(this.mSquareAspect_Click);
      // 
      // m43Aspect
      // 
      this.m43Aspect.Index = 2;
      this.m43Aspect.Text = "4:3";
      this.m43Aspect.Click += new System.EventHandler(this.m43Aspect_Click);
      // 
      // m169Aspect
      // 
      this.m169Aspect.Index = 3;
      this.m169Aspect.Text = "16:9";
      this.m169Aspect.Click += new System.EventHandler(this.m169Aspect_Click);
      // 
      // pPreviewVideo
      // 
      this.pPreviewVideo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
      | System.Windows.Forms.AnchorStyles.Left)
      | System.Windows.Forms.AnchorStyles.Right)));
      this.pPreviewVideo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.pPreviewVideo.Location = new System.Drawing.Point(8, 8);
      this.pPreviewVideo.Name = "pPreviewVideo";
      this.pPreviewVideo.Size = new System.Drawing.Size(480, 353);
      this.pPreviewVideo.TabIndex = 6;
      // 
      // CoolStreamBro
      // 
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.ClientSize = new System.Drawing.Size(496, 397);
      this.Controls.Add(this.pPreviewVideo);
      this.Controls.Add(this.bStop);
      this.Controls.Add(this.bStart);
      this.Controls.Add(this.lFilenameLabel);
      this.Controls.Add(this.tFilename);
      this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
      this.Menu = this.mainMenu;
      this.Name = "CoolStreamBro";
      this.Text = "CoolStreamBro";
      this.Load += new System.EventHandler(this.CaptureTest_Load);
      this.pPreviewVideo.ResumeLayout(false);
      this.ResumeLayout(false);
      this.PerformLayout();

    }
    #endregion


    [STAThread]
    static void Main()
    {
      AppDomain currentDomain = AppDomain.CurrentDomain;
      Application.Run(new CoolStreamBro());
    }

    private void bExit_Click(object sender, System.EventArgs e)
    {
      if (capture != null)
        capture.Stop();
      Application.Exit();
    }

    private void bStart_Click(object sender, System.EventArgs e)
    {
      try
      {
        if (capture == null)
          throw new ApplicationException("Please select input device.");
        if (!capture.Cued)
          capture.Filename = tFilename.Text;
        capture.Start();
        bStart.Enabled = false;
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void bStop_Click(object sender, System.EventArgs e)
    {
      try
      {
        if (capture == null)
          throw new ApplicationException("Please select input device.");
        capture.Stop();
        bStart.Enabled = true;
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message + "\n\n" + ex.ToString());
      }
    }

    /* Update all the menu options based on detected hardware */
    private void updateMenu()
    {
      MenuItem m;
      Filter f;
      Source s;
      Source current;
      PropertyPage p;
      Control oldPreviewWindow = null;

      /*
       * Cache the preview
       */
      if (capture != null)
      {
        oldPreviewWindow = capture.PreviewWindow;
        capture.PreviewWindow = null;
      }

      /*
       * Get all the video devices
       */
      Filter videoDevice = null;
      if (capture != null)
        videoDevice = capture.VideoDevice;
      mVideoDevices.MenuItems.Clear();
      m = new MenuItem("[empty]", new EventHandler(mVideoDevices_Click));
      m.Checked = (videoDevice == null);
      mVideoDevices.MenuItems.Add(m);
      for (int c = 0; c < filters.VideoInputDevices.Count; c++)
      {
        f = filters.VideoInputDevices[c];
        m = new MenuItem(f.Name, new EventHandler(mVideoDevices_Click));
        m.Checked = (videoDevice == f);
        mVideoDevices.MenuItems.Add(m);
      }
      mVideoDevices.Enabled = (filters.VideoInputDevices.Count > 0);

      /*
       * Get all the audio devices
       */
      Filter audioDevice = null;
      if (capture != null)
        audioDevice = capture.AudioDevice;
      mAudioDevices.MenuItems.Clear();
      m = new MenuItem("[empty]", new EventHandler(mAudioDevices_Click));
      m.Checked = (audioDevice == null);
      mAudioDevices.MenuItems.Add(m);
      for (int c = 0; c < filters.AudioInputDevices.Count; c++)
      {
        f = filters.AudioInputDevices[c];
        m = new MenuItem(f.Name, new EventHandler(mAudioDevices_Click));
        m.Checked = (audioDevice == f);
        mAudioDevices.MenuItems.Add(m);
      }
      mAudioDevices.Enabled = (filters.AudioInputDevices.Count > 0);

      /* 
       * Get all video codecs
       */
      try
      {
        mVideoCompressors.MenuItems.Clear();
        m = new MenuItem("[empty]", new EventHandler(mVideoCompressors_Click));
        m.Checked = (capture.VideoCompressor == null);
        mVideoCompressors.MenuItems.Add(m);
        for (int c = 0; c < filters.VideoCompressors.Count; c++)
        {
          f = filters.VideoCompressors[c];
          m = new MenuItem(f.Name, new EventHandler(mVideoCompressors_Click));
          m.Checked = (capture.VideoCompressor == f);
          mVideoCompressors.MenuItems.Add(m);
        }
        mVideoCompressors.Enabled = ((capture.VideoDevice != null) && (filters.VideoCompressors.Count > 0));
      }
      catch
      {
        mVideoCompressors.Enabled = false;
      }

      /* 
       * Get all the audio codecs
       */
      try
      {
        mAudioCompressors.MenuItems.Clear();
        m = new MenuItem("[empty]", new EventHandler(mAudioCompressors_Click));
        m.Checked = (capture.AudioCompressor == null);
        mAudioCompressors.MenuItems.Add(m);
        for (int c = 0; c < filters.AudioCompressors.Count; c++)
        {
          f = filters.AudioCompressors[c];
          m = new MenuItem(f.Name, new EventHandler(mAudioCompressors_Click));
          m.Checked = (capture.AudioCompressor == f);
          mAudioCompressors.MenuItems.Add(m);
        }
        mAudioCompressors.Enabled = ((capture.AudioDevice != null) && (filters.AudioCompressors.Count > 0));
      }
      catch
      {
        mAudioCompressors.Enabled = false;
      }

      /*
       * Get all video sources
       */
      try
      {
        mVideoSources.MenuItems.Clear();
        current = capture.VideoSource;
        for (int c = 0; c < capture.VideoSources.Count; c++)
        {
          s = capture.VideoSources[c];
          m = new MenuItem(s.Name, new EventHandler(mVideoSources_Click));
          m.Checked = (current == s);
          mVideoSources.MenuItems.Add(m);
        }
        mVideoSources.Enabled = (capture.VideoSources.Count > 0);
      }
      catch
      {
        mVideoSources.Enabled = false;
      }

      /*
       * Get all Audio Sources
       */
      try
      {
        mAudioSources.MenuItems.Clear();
        current = capture.AudioSource;
        for (int c = 0; c < capture.AudioSources.Count; c++)
        {
          s = capture.AudioSources[c];
          m = new MenuItem(s.Name, new EventHandler(mAudioSources_Click));
          m.Checked = (current == s);
          mAudioSources.MenuItems.Add(m);
        }
        mAudioSources.Enabled = (capture.AudioSources.Count > 0);
      }
      catch
      {
        mAudioSources.Enabled = false;
      }

      /* 
       * Add default frame rates
       */
      try
      {
        mFrameRates.MenuItems.Clear();
        int frameRate = (int)(capture.FrameRate * 1000);
        m = new MenuItem("15 fps", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 15000);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("24 fps (Film)", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 24000);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("25 fps (PAL)", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 25000);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("29.997 fps (NTSC)", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 29997);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("30 fps", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 30000);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("50 fps", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 50000);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("59.994 fps (Double NTSC)", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 59994);
        mFrameRates.MenuItems.Add(m);
        m = new MenuItem("60 fps", new EventHandler(mFrameRates_Click));
        m.Checked = (frameRate == 60000);
        mFrameRates.MenuItems.Add(m);
        mFrameRates.Enabled = true;
      }
      catch
      {
        mFrameRates.Enabled = false;
      }

      /* 
       * Add Default Capture Sizes
       */
      try
      {
        mFrameSizes.MenuItems.Clear();
        Size frameSize = capture.FrameSize;
        m = new MenuItem("160 x 120", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(160, 120));
        mFrameSizes.MenuItems.Add(m);
        m = new MenuItem("320 x 240", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(320, 240));
        mFrameSizes.MenuItems.Add(m);
        m = new MenuItem("426 x 240", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(426, 240));
        mFrameSizes.MenuItems.Add(m);
        m = new MenuItem("640 x 480", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(640, 480));
        mFrameSizes.MenuItems.Add(m);
        m = new MenuItem("1024 x 768", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(1024, 768));
        mFrameSizes.MenuItems.Add(m);
        m = new MenuItem("1280 x 720", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(1280, 720));
        mFrameSizes.MenuItems.Add(m);
        m = new MenuItem("1920 x 1080", new EventHandler(mFrameSizes_Click));
        m.Checked = (frameSize == new Size(1920, 1080));
        mFrameSizes.MenuItems.Add(m);
        mFrameSizes.Enabled = true;
      }
      catch
      {
        mFrameSizes.Enabled = false;
      }

      /* 
       * Add default audio channels
       */
      try
      {
        mAudioChannels.MenuItems.Clear();
        short audioChannels = capture.AudioChannels;
        m = new MenuItem("Mono", new EventHandler(mAudioChannels_Click));
        m.Checked = (audioChannels == 1);
        mAudioChannels.MenuItems.Add(m);
        m = new MenuItem("Stereo", new EventHandler(mAudioChannels_Click));
        m.Checked = (audioChannels == 2);
        mAudioChannels.MenuItems.Add(m);
        mAudioChannels.Enabled = true;
      }
      catch
      {
        mAudioChannels.Enabled = false;
      }

      /*
       * Add Audio sampling rates
       */
      try
      {
        mAudioSamplingRate.MenuItems.Clear();
        int samplingRate = capture.AudioSamplingRate;
        m = new MenuItem("8 kHz", new EventHandler(mAudioSamplingRate_Click));
        m.Checked = (samplingRate == 8000);
        mAudioSamplingRate.MenuItems.Add(m);
        m = new MenuItem("11.025 kHz", new EventHandler(mAudioSamplingRate_Click));
        m.Checked = (capture.AudioSamplingRate == 11025);
        mAudioSamplingRate.MenuItems.Add(m);
        m = new MenuItem("22.05 kHz", new EventHandler(mAudioSamplingRate_Click));
        m.Checked = (capture.AudioSamplingRate == 22050);
        mAudioSamplingRate.MenuItems.Add(m);
        m = new MenuItem("44.1 kHz", new EventHandler(mAudioSamplingRate_Click));
        m.Checked = (capture.AudioSamplingRate == 44100);
        mAudioSamplingRate.MenuItems.Add(m);
        mAudioSamplingRate.Enabled = true;
      }
      catch
      {
        mAudioSamplingRate.Enabled = false;
      }

      /*
       * Add audio sample qualities
       */
      try
      {
        mAudioSampleSizes.MenuItems.Clear();
        short sampleSize = capture.AudioSampleSize;
        m = new MenuItem("8 bit", new EventHandler(mAudioSampleSizes_Click));
        m.Checked = (sampleSize == 8);
        mAudioSampleSizes.MenuItems.Add(m);
        m = new MenuItem("16 bit", new EventHandler(mAudioSampleSizes_Click));
        m.Checked = (sampleSize == 16);
        mAudioSampleSizes.MenuItems.Add(m);
        mAudioSampleSizes.Enabled = true;
      }
      catch
      {
        mAudioSampleSizes.Enabled = false;
      }

      /*
       * Add the OLE property pages
       */
      try
      {
        mPropertyPages.MenuItems.Clear();
        for (int c = 0; c < capture.PropertyPages.Count; c++)
        {
          p = capture.PropertyPages[c];
          m = new MenuItem(p.Name + "...", new EventHandler(mPropertyPages_Click));
          mPropertyPages.MenuItems.Add(m);
        }
        mPropertyPages.Enabled = (capture.PropertyPages.Count > 0);
      }
      catch
      {
        mPropertyPages.Enabled = false;
      }

      /*
       * Add any TV tuner channels
       */
      try
      {
        mChannel.MenuItems.Clear();
        int channel = capture.Tuner.Channel;
        for (int c = 1; c <= 25; c++)
        {
          m = new MenuItem(c.ToString(), new EventHandler(mChannel_Click));
          m.Checked = (channel == c);
          mChannel.MenuItems.Add(m);
        }
        mChannel.Enabled = true;
      }
      catch
      {
        mChannel.Enabled = false;
      }

      /*
       * Add any TV tuner inputs
       */
      try
      {
        mInputType.MenuItems.Clear();
        m = new MenuItem(TunerInputType.Cable.ToString(), new EventHandler(mInputType_Click));
        m.Checked = (capture.Tuner.InputType == TunerInputType.Cable);
        mInputType.MenuItems.Add(m);
        m = new MenuItem(TunerInputType.Antenna.ToString(), new EventHandler(mInputType_Click));
        m.Checked = (capture.Tuner.InputType == TunerInputType.Antenna);
        mInputType.MenuItems.Add(m);
        mInputType.Enabled = true;
      }
      catch
      {
        mInputType.Enabled = false;
      }

      /*
       * Enable or disable device capabilities
       */
      mVideoCaps.Enabled = ((capture != null) && (capture.VideoCaps != null));
      mAudioCaps.Enabled = ((capture != null) && (capture.AudioCaps != null));

      /*
       * Enable preview
       */
      mPreviewEnable.Checked = (oldPreviewWindow != null);
      mPreviewEnable.Enabled = (capture != null);

      /*
       * Restore the preview
       */
      if (capture != null)
        capture.PreviewWindow = oldPreviewWindow;
    }

    private void mVideoDevices_Click(object sender, System.EventArgs e)
    {
      try
      {
        /*
         * Dispose everything and start from scratch
         */
        Filter videoDevice = null;
        Filter audioDevice = null;
        if (capture != null)
        {
          videoDevice = capture.VideoDevice;
          audioDevice = capture.AudioDevice;
          capture.Dispose();
          capture = null;
        }

        /*
         * Get new video device
         */
        MenuItem m = sender as MenuItem;
        videoDevice = (m.Index > 0 ? filters.VideoInputDevices[m.Index - 1] : null);

        /*
         * Create capture object
         */
        if ((videoDevice != null) || (audioDevice != null))
        {
          capture = new Capture(videoDevice, audioDevice);
          capture.CaptureComplete += new EventHandler(OnCaptureComplete);
        }

        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Unsupported Video Device?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mAudioDevices_Click(object sender, System.EventArgs e)
    {
      try
      {
        /*
         * Dispose everything and try again
         */
        Filter videoDevice = null;
        Filter audioDevice = null;
        if (capture != null)
        {
          videoDevice = capture.VideoDevice;
          audioDevice = capture.AudioDevice;
          capture.Dispose();
          capture = null;
        }

        /* 
         * Get new audio device
         */
        MenuItem m = sender as MenuItem;
        audioDevice = (m.Index > 0 ? filters.AudioInputDevices[m.Index - 1] : null);

        /*
         * Create capture object
         */
        if ((videoDevice != null) || (audioDevice != null))
        {
          capture = new Capture(videoDevice, audioDevice);
          capture.CaptureComplete += new EventHandler(OnCaptureComplete);
        }

        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Unsupported Audio Device?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mVideoCompressors_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.VideoCompressor = (m.Index > 0 ? filters.VideoCompressors[m.Index - 1] : null);
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Video codec not supported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }

    }

    private void mAudioCompressors_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.AudioCompressor = (m.Index > 0 ? filters.AudioCompressors[m.Index - 1] : null);
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Audio codec not supported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mVideoSources_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.VideoSource = capture.VideoSources[m.Index];
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Failure setting video source.\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mAudioSources_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.AudioSource = capture.AudioSources[m.Index];
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Failure setting audio source.\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }


    private void mExit_Click(object sender, System.EventArgs e)
    {
      /* quit */
      if (capture != null)
        capture.Stop();
      Application.Exit();
    }

    private void mFrameSizes_Click(object sender, System.EventArgs e)
    {
      try
      {
        bool preview = (capture.PreviewWindow != null);
        capture.PreviewWindow = null;

        MenuItem m = sender as MenuItem;
        string[] s = m.Text.Split('x');
        Size size = new Size(int.Parse(s[0]), int.Parse(s[1]));
        capture.FrameSize = size;

        updateMenu();

        capture.PreviewWindow = (preview ? pPreviewVideo : null);
      }
      catch (Exception ex)
      {
        MessageBox.Show("Resolution not supported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mFrameRates_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        string[] s = m.Text.Split(' ');
        capture.FrameRate = double.Parse(s[0]);
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("FPS not supported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }


    private void mAudioChannels_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.AudioChannels = (short)Math.Pow(2, m.Index);
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Audio channel selection unsupported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mAudioSamplingRate_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        string[] s = m.Text.Split(' ');
        int samplingRate = (int)(double.Parse(s[0]) * 1000);
        capture.AudioSamplingRate = samplingRate;
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Sampling rate unsupported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mAudioSampleSizes_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        string[] s = m.Text.Split(' ');
        short sampleSize = short.Parse(s[0]);
        capture.AudioSampleSize = sampleSize;
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Audio depth unsupported?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mPreview_Click(object sender, System.EventArgs e)
    {
      try
      {
        if (capture.PreviewWindow == null)
        {
          capture.PreviewWindow = pPreviewVideo;
          mPreviewEnable.Checked = true;
        }
        else
        {
          capture.PreviewWindow = null;
          mPreviewEnable.Checked = false;

        }
      }
      catch (Exception ex)
      {
        MessageBox.Show("Something broke toggling preview.\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mPropertyPages_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.PropertyPages[m.Index].Show(this);
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Something broke showing the property page.\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mChannel_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.Tuner.Channel = m.Index + 1;
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Channel change failed?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mInputType_Click(object sender, System.EventArgs e)
    {
      try
      {
        MenuItem m = sender as MenuItem;
        capture.Tuner.InputType = (TunerInputType)m.Index;
        updateMenu();
      }
      catch (Exception ex)
      {
        MessageBox.Show("Tuner input type change failed?\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mVideoCaps_Click(object sender, System.EventArgs e)
    {
      try
      {
        string s;
        s = String.Format(
            "Video Device Capabilities\n" +
            "=========================\n\n" +
            "Input Size:\t\t{0} x {1}\n" +
            "Min Frame Size:\t\t{2} x {3}\n" +
            "Max Frame Size:\t\t{4} x {5}\n\n" +
            "Frame Size Granularity X:\t{6}\n" +
            "Frame Size Granularity Y:\t{7}\n\n" +
            "Min Frame Rate:\t\t{8:0.0000} fps\n" +
            "Max Frame Rate:\t\t{9:0.0000} fps\n",
            capture.VideoCaps.InputSize.Width, capture.VideoCaps.InputSize.Height,
            capture.VideoCaps.MinFrameSize.Width, capture.VideoCaps.MinFrameSize.Height,
            capture.VideoCaps.MaxFrameSize.Width, capture.VideoCaps.MaxFrameSize.Height,
            capture.VideoCaps.FrameSizeGranularityX,
            capture.VideoCaps.FrameSizeGranularityY,
            capture.VideoCaps.MinFrameRate,
            capture.VideoCaps.MaxFrameRate);
        MessageBox.Show(s);

      }
      catch (Exception ex)
      {
        MessageBox.Show("Video Device Capabilities\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mAudioCaps_Click(object sender, System.EventArgs e)
    {
      try
      {
        string s;
        s = String.Format(
            "Audio Device Capabilities\n" +
            "=========================\n\n" +
            "Min Channels:\t\t{0}\n" +
            "Max Channels:\t\t{1}\n" +
            "Channels Granularity:\t{2}\n\n" +
            "Min Sample Size:\t\t{3}\n" +
            "Max Sample Size:\t\t{4}\n" +
            "Sample Size Granularity:\t{5}\n\n" +
            "Min Sampling Rate:\t\t{6}\n" +
            "Max Sampling Rate:\t\t{7}\n" +
            "Sampling Rate Granularity:\t{8}\n",
            capture.AudioCaps.MinimumChannels,
            capture.AudioCaps.MaximumChannels,
            capture.AudioCaps.ChannelsGranularity,
            capture.AudioCaps.MinimumSampleSize,
            capture.AudioCaps.MaximumSampleSize,
            capture.AudioCaps.SampleSizeGranularity,
            capture.AudioCaps.MinimumSamplingRate,
            capture.AudioCaps.MaximumSamplingRate,
            capture.AudioCaps.SamplingRateGranularity);
        MessageBox.Show(s);

      }
      catch (Exception ex)
      {
        MessageBox.Show("Audio Device Capabilities\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void OnCaptureComplete(object sender, EventArgs e)
    {
      Debug.WriteLine("Capture complete.");
    }

    private void CaptureTest_Load(object sender, EventArgs e)
    {

    }

    private void txtFilename_TextChanged(object sender, EventArgs e)
    {

    }

    private void mPreviewAudio_Click(object sender, EventArgs e)
    {
      /* enable monitoring audio */
      try
      {
        if (capture.wantAudioPreview == false)
        {
          capture.wantAudioPreview = true;
          mPreviewAudio.Checked = true;
        }
        else
        {
          capture.wantAudioPreview = false;
          mPreviewAudio.Checked = false;
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show("Toggling audio monitoring during preview failed.\n\n" + ex.Message + "\n\n" + ex.ToString());
      }
    }

    private void mFreeAspect_Click(object sender, EventArgs e)
    {
      fixed_aspect = false;
      mFreeAspect.Checked = true;
      mSquareAspect.Checked = false;
      m43Aspect.Checked = false;
      m169Aspect.Checked = false;
    }

    private void mSquareAspect_Click(object sender, EventArgs e)
    {
      fixed_aspect = true;
      aspect = 1f;
      mFreeAspect.Checked = false;
      mSquareAspect.Checked = true;
      m43Aspect.Checked = false;
      m169Aspect.Checked = false;
    }

    private void m43Aspect_Click(object sender, EventArgs e)
    {
      fixed_aspect = true;
      aspect = 1.333f;
      mFreeAspect.Checked = false;
      mSquareAspect.Checked = false;
      m43Aspect.Checked = true;
      m169Aspect.Checked = false;
    }

    private void m169Aspect_Click(object sender, EventArgs e)
    {
      fixed_aspect = true;
      aspect = 1.777f;
      mFreeAspect.Checked = false;
      mSquareAspect.Checked = false;
      m43Aspect.Checked = false;
      m169Aspect.Checked = true;
    }
  }
}
