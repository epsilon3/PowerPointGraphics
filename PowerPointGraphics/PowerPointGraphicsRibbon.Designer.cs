using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace PowerPointGraphics
{
    partial class PowerPointGraphics : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PowerPointGraphics()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PowerPointGraphics));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tabAppBuilderGraphics = this.Factory.CreateRibbonTab();
            this.grpPowerPointGraphicsTools = this.Factory.CreateRibbonGroup();
            this.butLoadPictures = this.Factory.CreateRibbonButton();
            this.butSavePictures = this.Factory.CreateRibbonButton();
            this.butProcessPicturesTop = this.Factory.CreateRibbonButton();
            this.butProcessPicturesLeft = this.Factory.CreateRibbonButton();
            this.butProcessPicturesWidth = this.Factory.CreateRibbonButton();
            this.butProcessPicturesHeight = this.Factory.CreateRibbonButton();
            this.butClearSlides = this.Factory.CreateRibbonButton();
            this.grpGraphicsSettings = this.Factory.CreateRibbonGroup();
            this.chkPositionPicture = this.Factory.CreateRibbonCheckBox();
            this.chkCropPicture = this.Factory.CreateRibbonCheckBox();
            this.comboGraphicsExport = this.Factory.CreateRibbonComboBox();
            this.editPictureOriginalWidthPix = this.Factory.CreateRibbonEditBox();
            this.editPictureOriginalHeightPix = this.Factory.CreateRibbonEditBox();
            this.editPicturePositionLeftPix = this.Factory.CreateRibbonEditBox();
            this.editPicturePositionTopPix = this.Factory.CreateRibbonEditBox();
            this.editPicturePositionWidthPix = this.Factory.CreateRibbonEditBox();
            this.editPicturePositionHeightPix = this.Factory.CreateRibbonEditBox();
            this.editPictureCropLeftPix = this.Factory.CreateRibbonEditBox();
            this.editPictureCropTopPix = this.Factory.CreateRibbonEditBox();
            this.editPictureCropWidthPix = this.Factory.CreateRibbonEditBox();
            this.editPictureCropHeightPix = this.Factory.CreateRibbonEditBox();
            this.grpSettings = this.Factory.CreateRibbonGroup();
            this.butLoadDefaults = this.Factory.CreateRibbonButton();
            this.butSaveDefaults = this.Factory.CreateRibbonButton();
            this.butSetDefaults = this.Factory.CreateRibbonButton();
            this.editDefaultsFileName = this.Factory.CreateRibbonEditBox();
            this.butBrowseBaseFolder = this.Factory.CreateRibbonButton();
            this.butBrowseGraphicsFolder = this.Factory.CreateRibbonButton();
            this.editBaseFolder = this.Factory.CreateRibbonEditBox();
            this.editGraphicsFolder = this.Factory.CreateRibbonEditBox();
            this.editFileExtension = this.Factory.CreateRibbonEditBox();
            this.butProcessPicturesCrop = this.Factory.CreateRibbonButton();
            this.tabAppBuilderGraphics.SuspendLayout();
            this.grpPowerPointGraphicsTools.SuspendLayout();
            this.grpGraphicsSettings.SuspendLayout();
            this.grpSettings.SuspendLayout();
            // 
            // tabAppBuilderGraphics
            // 
            this.tabAppBuilderGraphics.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabAppBuilderGraphics.Groups.Add(this.grpPowerPointGraphicsTools);
            this.tabAppBuilderGraphics.Groups.Add(this.grpGraphicsSettings);
            this.tabAppBuilderGraphics.Groups.Add(this.grpSettings);
            this.tabAppBuilderGraphics.Label = "POWERPOINTGRAPHICS";
            this.tabAppBuilderGraphics.Name = "tabAppBuilderGraphics";
            // 
            // grpPowerPointGraphicsTools
            // 
            this.grpPowerPointGraphicsTools.Items.Add(this.butLoadPictures);
            this.grpPowerPointGraphicsTools.Items.Add(this.butSavePictures);
            this.grpPowerPointGraphicsTools.Items.Add(this.butProcessPicturesCrop);
            this.grpPowerPointGraphicsTools.Items.Add(this.butProcessPicturesTop);
            this.grpPowerPointGraphicsTools.Items.Add(this.butProcessPicturesLeft);
            this.grpPowerPointGraphicsTools.Items.Add(this.butProcessPicturesWidth);
            this.grpPowerPointGraphicsTools.Items.Add(this.butProcessPicturesHeight);
            this.grpPowerPointGraphicsTools.Items.Add(this.butClearSlides);
            this.grpPowerPointGraphicsTools.Label = "PowerPointGraphics Tools";
            this.grpPowerPointGraphicsTools.Name = "grpPowerPointGraphicsTools";
            // 
            // butLoadPictures
            // 
            this.butLoadPictures.Image = ((System.Drawing.Image)(resources.GetObject("butLoadPictures.Image")));
            this.butLoadPictures.Label = "Load Pictures";
            this.butLoadPictures.Name = "butLoadPictures";
            this.butLoadPictures.ScreenTip = "Load Pictures";
            this.butLoadPictures.ShowImage = true;
            this.butLoadPictures.SuperTip = "Load Pictures";
            this.butLoadPictures.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butLoadPictures_Click);
            // 
            // butSavePictures
            // 
            this.butSavePictures.Image = ((System.Drawing.Image)(resources.GetObject("butSavePictures.Image")));
            this.butSavePictures.Label = "Save Pictures";
            this.butSavePictures.Name = "butSavePictures";
            this.butSavePictures.ScreenTip = "Save Pictures";
            this.butSavePictures.ShowImage = true;
            this.butSavePictures.SuperTip = "Save Pictures";
            this.butSavePictures.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butSavePictures_Click);
            // 
            // butProcessPicturesTop
            // 
            this.butProcessPicturesTop.Image = ((System.Drawing.Image)(resources.GetObject("butProcessPicturesTop.Image")));
            this.butProcessPicturesTop.Label = "Process Pictures (Top)";
            this.butProcessPicturesTop.Name = "butProcessPicturesTop";
            this.butProcessPicturesTop.ScreenTip = "Process Pictures (Top)";
            this.butProcessPicturesTop.ShowImage = true;
            this.butProcessPicturesTop.SuperTip = "Process Pictures (Top)";
            this.butProcessPicturesTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butProcessPicturesTop_Click);
            // 
            // butProcessPicturesLeft
            // 
            this.butProcessPicturesLeft.Image = ((System.Drawing.Image)(resources.GetObject("butProcessPicturesLeft.Image")));
            this.butProcessPicturesLeft.Label = "Process Pictures (Left)";
            this.butProcessPicturesLeft.Name = "butProcessPicturesLeft";
            this.butProcessPicturesLeft.ScreenTip = "Process Pictures (Left)";
            this.butProcessPicturesLeft.ShowImage = true;
            this.butProcessPicturesLeft.SuperTip = "Process Pictures (Left)";
            this.butProcessPicturesLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butProcessPicturesLeft_Click);
            // 
            // butProcessPicturesWidth
            // 
            this.butProcessPicturesWidth.Image = ((System.Drawing.Image)(resources.GetObject("butProcessPicturesWidth.Image")));
            this.butProcessPicturesWidth.Label = "Process Pictures (Width)";
            this.butProcessPicturesWidth.Name = "butProcessPicturesWidth";
            this.butProcessPicturesWidth.ScreenTip = "Process Pictures (Width)";
            this.butProcessPicturesWidth.ShowImage = true;
            this.butProcessPicturesWidth.SuperTip = "Process Pictures (Width)";
            this.butProcessPicturesWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butProcessPicturesWidth_Click);
            // 
            // butProcessPicturesHeight
            // 
            this.butProcessPicturesHeight.Image = ((System.Drawing.Image)(resources.GetObject("butProcessPicturesHeight.Image")));
            this.butProcessPicturesHeight.Label = "Process Pictures (Height)";
            this.butProcessPicturesHeight.Name = "butProcessPicturesHeight";
            this.butProcessPicturesHeight.ScreenTip = "Process Pictures (Height)";
            this.butProcessPicturesHeight.ShowImage = true;
            this.butProcessPicturesHeight.SuperTip = "Process Pictures (Height)";
            this.butProcessPicturesHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butProcessPicturesHeight_Click);
            // 
            // butClearSlides
            // 
            this.butClearSlides.Image = ((System.Drawing.Image)(resources.GetObject("butClearSlides.Image")));
            this.butClearSlides.Label = "Clear Slides";
            this.butClearSlides.Name = "butClearSlides";
            this.butClearSlides.ScreenTip = "Clear Slides";
            this.butClearSlides.ShowImage = true;
            this.butClearSlides.SuperTip = "Clear Slides";
            this.butClearSlides.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butClearSlides_Click);
            // 
            // grpGraphicsSettings
            // 
            this.grpGraphicsSettings.Items.Add(this.chkPositionPicture);
            this.grpGraphicsSettings.Items.Add(this.chkCropPicture);
            this.grpGraphicsSettings.Items.Add(this.comboGraphicsExport);
            this.grpGraphicsSettings.Items.Add(this.editPictureOriginalWidthPix);
            this.grpGraphicsSettings.Items.Add(this.editPictureOriginalHeightPix);
            this.grpGraphicsSettings.Items.Add(this.editPicturePositionLeftPix);
            this.grpGraphicsSettings.Items.Add(this.editPicturePositionTopPix);
            this.grpGraphicsSettings.Items.Add(this.editPicturePositionWidthPix);
            this.grpGraphicsSettings.Items.Add(this.editPicturePositionHeightPix);
            this.grpGraphicsSettings.Items.Add(this.editPictureCropLeftPix);
            this.grpGraphicsSettings.Items.Add(this.editPictureCropTopPix);
            this.grpGraphicsSettings.Items.Add(this.editPictureCropWidthPix);
            this.grpGraphicsSettings.Items.Add(this.editPictureCropHeightPix);
            this.grpGraphicsSettings.Label = "PowerPointGraphics Settings";
            this.grpGraphicsSettings.Name = "grpGraphicsSettings";
            // 
            // chkPositionPicture
            // 
            this.chkPositionPicture.Checked = true;
            this.chkPositionPicture.Label = "Position Picture (On Load)";
            this.chkPositionPicture.Name = "chkPositionPicture";
            this.chkPositionPicture.ScreenTip = "Position Picture (On Load)";
            this.chkPositionPicture.SuperTip = "Position Picture (On Load)";
            this.chkPositionPicture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkPositionPicture_Click);
            // 
            // chkCropPicture
            // 
            this.chkCropPicture.Label = "Crop Picture (On Load)";
            this.chkCropPicture.Name = "chkCropPicture";
            this.chkCropPicture.ScreenTip = "Crop Picture (On Load)";
            this.chkCropPicture.SuperTip = "Crop Picture (On Load)";
            this.chkCropPicture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkCropPicture_Click);
            // 
            // comboGraphicsExport
            // 
            ribbonDropDownItemImpl1.Label = "Export Picture(s) to Graphic(s) (Excludes Slide Background)";
            ribbonDropDownItemImpl1.ScreenTip = "Export Picture(s) to Graphic(s) (Excludes Slide Background)";
            ribbonDropDownItemImpl1.SuperTip = "Export Picture(s) to Graphic(s) (Excludes Slide Background)";
            ribbonDropDownItemImpl2.Label = "Export Slide(s) to Graphic(s) (Includes Slide Background)";
            ribbonDropDownItemImpl2.ScreenTip = "Export Slide(s) to Graphic(s) (Includes Slide Background)";
            ribbonDropDownItemImpl2.SuperTip = "Export Slide(s) to Graphic(s) (Includes Slide Background)";
            this.comboGraphicsExport.Items.Add(ribbonDropDownItemImpl1);
            this.comboGraphicsExport.Items.Add(ribbonDropDownItemImpl2);
            this.comboGraphicsExport.Label = "Graphics Export Mode:";
            this.comboGraphicsExport.Name = "comboGraphicsExport";
            this.comboGraphicsExport.ScreenTip = "Graphics Export Mode";
            this.comboGraphicsExport.SizeString = "........................................";
            this.comboGraphicsExport.SuperTip = "Graphics Export Mode";
            this.comboGraphicsExport.Text = "Export Picture(s) to Graphic(s) (Excludes Slide Background)";
            this.comboGraphicsExport.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboGraphicsExport_TextChanged);
            // 
            // editPictureOriginalWidthPix
            // 
            this.editPictureOriginalWidthPix.Label = "Picture Original Width (Pixels):";
            this.editPictureOriginalWidthPix.Name = "editPictureOriginalWidthPix";
            this.editPictureOriginalWidthPix.ScreenTip = "Picture Original Width (Pixels)";
            this.editPictureOriginalWidthPix.SizeString = "..........";
            this.editPictureOriginalWidthPix.SuperTip = "Picture Original Width (Pixels)";
            this.editPictureOriginalWidthPix.Text = "1122";
            this.editPictureOriginalWidthPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPictureOriginalWidthPix_TextChanged);
            // 
            // editPictureOriginalHeightPix
            // 
            this.editPictureOriginalHeightPix.Label = "Picture Original Height (Pixels):";
            this.editPictureOriginalHeightPix.Name = "editPictureOriginalHeightPix";
            this.editPictureOriginalHeightPix.ScreenTip = "Picture Original Height (Pixels)";
            this.editPictureOriginalHeightPix.SizeString = "..........";
            this.editPictureOriginalHeightPix.SuperTip = "Picture Original Height (Pixels)";
            this.editPictureOriginalHeightPix.Text = "744";
            this.editPictureOriginalHeightPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPictureOriginalHeightPix_TextChanged);
            // 
            // editPicturePositionLeftPix
            // 
            this.editPicturePositionLeftPix.Label = "Picture Position Left (Pixels):";
            this.editPicturePositionLeftPix.Name = "editPicturePositionLeftPix";
            this.editPicturePositionLeftPix.ScreenTip = "Picture Position Left (Pixels)";
            this.editPicturePositionLeftPix.SizeString = "..........";
            this.editPicturePositionLeftPix.SuperTip = "Picture Position Left (Pixels)";
            this.editPicturePositionLeftPix.Text = "153";
            this.editPicturePositionLeftPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPicturePositionLeftPix_TextChanged);
            // 
            // editPicturePositionTopPix
            // 
            this.editPicturePositionTopPix.Label = "Picture Position Top (Pixels):";
            this.editPicturePositionTopPix.Name = "editPicturePositionTopPix";
            this.editPicturePositionTopPix.ScreenTip = "Picture Position Top (Pixels)";
            this.editPicturePositionTopPix.SizeString = "..........";
            this.editPicturePositionTopPix.SuperTip = "Picture Position Top (Pixels)";
            this.editPicturePositionTopPix.Text = "0";
            this.editPicturePositionTopPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPicturePositionTopPix_TextChanged);
            // 
            // editPicturePositionWidthPix
            // 
            this.editPicturePositionWidthPix.Label = "Picture Position Width (Pixels):";
            this.editPicturePositionWidthPix.Name = "editPicturePositionWidthPix";
            this.editPicturePositionWidthPix.ScreenTip = "Picture Position Width (Pixels)";
            this.editPicturePositionWidthPix.SizeString = "..........";
            this.editPicturePositionWidthPix.SuperTip = "Picture Position Width (Pixels)";
            this.editPicturePositionWidthPix.Text = "600";
            this.editPicturePositionWidthPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPicturePositionWidthPix_TextChanged);
            // 
            // editPicturePositionHeightPix
            // 
            this.editPicturePositionHeightPix.Label = "Picture Position Height (Pixels):";
            this.editPicturePositionHeightPix.Name = "editPicturePositionHeightPix";
            this.editPicturePositionHeightPix.ScreenTip = "Picture Position Height (Pixels)";
            this.editPicturePositionHeightPix.SizeString = "..........";
            this.editPicturePositionHeightPix.SuperTip = "Picture Position Height (Pixels)";
            this.editPicturePositionHeightPix.Text = "600";
            this.editPicturePositionHeightPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPicturePositionHeightPix_TextChanged);
            // 
            // editPictureCropLeftPix
            // 
            this.editPictureCropLeftPix.Label = "Picture Crop Left (Pixels):";
            this.editPictureCropLeftPix.Name = "editPictureCropLeftPix";
            this.editPictureCropLeftPix.ScreenTip = "Picture Crop Left (Pixels)";
            this.editPictureCropLeftPix.SizeString = "..........";
            this.editPictureCropLeftPix.SuperTip = "Picture Crop Left (Pixels)";
            this.editPictureCropLeftPix.Text = "144";
            this.editPictureCropLeftPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPictureCropLeftPix_TextChanged);
            // 
            // editPictureCropTopPix
            // 
            this.editPictureCropTopPix.Label = "Picture Crop Top (Pixels):";
            this.editPictureCropTopPix.Name = "editPictureCropTopPix";
            this.editPictureCropTopPix.ScreenTip = "Picture Crop Top (Pixels)";
            this.editPictureCropTopPix.SizeString = "..........";
            this.editPictureCropTopPix.SuperTip = "Picture Crop Top (Pixels)";
            this.editPictureCropTopPix.Text = "144";
            this.editPictureCropTopPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPictureCropTopPix_TextChanged);
            // 
            // editPictureCropWidthPix
            // 
            this.editPictureCropWidthPix.Label = "Picture Crop Width (Pixels):";
            this.editPictureCropWidthPix.Name = "editPictureCropWidthPix";
            this.editPictureCropWidthPix.ScreenTip = "Picture Crop Width (Pixels)";
            this.editPictureCropWidthPix.SizeString = "..........";
            this.editPictureCropWidthPix.SuperTip = "Picture Crop Width (Pixels)";
            this.editPictureCropWidthPix.Text = "336";
            this.editPictureCropWidthPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPictureCropWidthPix_TextChanged);
            // 
            // editPictureCropHeightPix
            // 
            this.editPictureCropHeightPix.Label = "Picture Crop Height (Pixels):";
            this.editPictureCropHeightPix.Name = "editPictureCropHeightPix";
            this.editPictureCropHeightPix.ScreenTip = "Picture Crop Height (Pixels)";
            this.editPictureCropHeightPix.SizeString = "..........";
            this.editPictureCropHeightPix.SuperTip = "Picture Crop Height (Pixels)";
            this.editPictureCropHeightPix.Text = "336";
            this.editPictureCropHeightPix.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPictureCropHeightPix_TextChanged);
            // 
            // grpSettings
            // 
            this.grpSettings.Items.Add(this.butLoadDefaults);
            this.grpSettings.Items.Add(this.butSaveDefaults);
            this.grpSettings.Items.Add(this.butSetDefaults);
            this.grpSettings.Items.Add(this.editDefaultsFileName);
            this.grpSettings.Items.Add(this.butBrowseBaseFolder);
            this.grpSettings.Items.Add(this.butBrowseGraphicsFolder);
            this.grpSettings.Items.Add(this.editBaseFolder);
            this.grpSettings.Items.Add(this.editGraphicsFolder);
            this.grpSettings.Items.Add(this.editFileExtension);
            this.grpSettings.Label = "PowerPointGraphics General Settings";
            this.grpSettings.Name = "grpSettings";
            // 
            // butLoadDefaults
            // 
            this.butLoadDefaults.Image = ((System.Drawing.Image)(resources.GetObject("butLoadDefaults.Image")));
            this.butLoadDefaults.Label = "Load Defaults";
            this.butLoadDefaults.Name = "butLoadDefaults";
            this.butLoadDefaults.ScreenTip = "Load Defaults";
            this.butLoadDefaults.ShowImage = true;
            this.butLoadDefaults.SuperTip = "Load Defaults";
            this.butLoadDefaults.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butLoadDefaults_Click);
            // 
            // butSaveDefaults
            // 
            this.butSaveDefaults.Image = ((System.Drawing.Image)(resources.GetObject("butSaveDefaults.Image")));
            this.butSaveDefaults.Label = "Save Defaults";
            this.butSaveDefaults.Name = "butSaveDefaults";
            this.butSaveDefaults.ScreenTip = "Save Defaults";
            this.butSaveDefaults.ShowImage = true;
            this.butSaveDefaults.SuperTip = "Save Defaults";
            this.butSaveDefaults.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butSaveDefaults_Click);
            // 
            // butSetDefaults
            // 
            this.butSetDefaults.Image = ((System.Drawing.Image)(resources.GetObject("butSetDefaults.Image")));
            this.butSetDefaults.Label = "Set Defaults";
            this.butSetDefaults.Name = "butSetDefaults";
            this.butSetDefaults.ScreenTip = "Set Defaults";
            this.butSetDefaults.ShowImage = true;
            this.butSetDefaults.SuperTip = "Set Defaults";
            this.butSetDefaults.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butSetDefaults_Click);
            // 
            // editDefaultsFileName
            // 
            this.editDefaultsFileName.Label = "Defaults File Name:";
            this.editDefaultsFileName.Name = "editDefaultsFileName";
            this.editDefaultsFileName.ScreenTip = "Defaults File Name";
            this.editDefaultsFileName.SizeString = "............................................................";
            this.editDefaultsFileName.SuperTip = "Defaults File Name";
            this.editDefaultsFileName.Text = "PowerPointGraphics_Defaults.txt";
            // 
            // butBrowseBaseFolder
            // 
            this.butBrowseBaseFolder.Image = ((System.Drawing.Image)(resources.GetObject("butBrowseBaseFolder.Image")));
            this.butBrowseBaseFolder.Label = "Browse Base Folder";
            this.butBrowseBaseFolder.Name = "butBrowseBaseFolder";
            this.butBrowseBaseFolder.ScreenTip = "Browse Base Folder";
            this.butBrowseBaseFolder.ShowImage = true;
            this.butBrowseBaseFolder.SuperTip = "Browse Base Folder";
            this.butBrowseBaseFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butBrowseBaseFolder_Click);
            // 
            // butBrowseGraphicsFolder
            // 
            this.butBrowseGraphicsFolder.Image = ((System.Drawing.Image)(resources.GetObject("butBrowseGraphicsFolder.Image")));
            this.butBrowseGraphicsFolder.Label = "Browse Graphics Folder";
            this.butBrowseGraphicsFolder.Name = "butBrowseGraphicsFolder";
            this.butBrowseGraphicsFolder.ScreenTip = "Browse Graphics Folder";
            this.butBrowseGraphicsFolder.ShowImage = true;
            this.butBrowseGraphicsFolder.SuperTip = "Browse Graphics Folder";
            this.butBrowseGraphicsFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butBrowseGraphicsFolder_Click);
            // 
            // editBaseFolder
            // 
            this.editBaseFolder.Label = "Base Folder:";
            this.editBaseFolder.Name = "editBaseFolder";
            this.editBaseFolder.ScreenTip = "Base Folder";
            this.editBaseFolder.SizeString = "............................................................";
            this.editBaseFolder.SuperTip = "Base Folder";
            this.editBaseFolder.Text = "C:\\HuffStuff\\STRONG_Lab\\Schemas";
            // 
            // editGraphicsFolder
            // 
            this.editGraphicsFolder.Label = "Graphics Folder:";
            this.editGraphicsFolder.Name = "editGraphicsFolder";
            this.editGraphicsFolder.ScreenTip = "Graphics Folder";
            this.editGraphicsFolder.SizeString = "............................................................";
            this.editGraphicsFolder.SuperTip = "Graphics Folder";
            this.editGraphicsFolder.Text = "C:\\Users\\Stephen Huff\\AppData\\Local\\Packages\\AFFECTS_6m8grht3agkqy\\LocalState";
            // 
            // editFileExtension
            // 
            this.editFileExtension.Label = "File Extension:";
            this.editFileExtension.Name = "editFileExtension";
            this.editFileExtension.ScreenTip = "File Extension";
            this.editFileExtension.SizeString = "..........";
            this.editFileExtension.SuperTip = "File Extension";
            this.editFileExtension.Text = "png";
            // 
            // butProcessPicturesCrop
            // 
            this.butProcessPicturesCrop.Image = ((System.Drawing.Image)(resources.GetObject("butProcessPicturesCrop.Image")));
            this.butProcessPicturesCrop.Label = "Process Pictures (Crop)";
            this.butProcessPicturesCrop.Name = "butProcessPicturesCrop";
            this.butProcessPicturesCrop.ScreenTip = "Process Pictures (Crop)";
            this.butProcessPicturesCrop.ShowImage = true;
            this.butProcessPicturesCrop.SuperTip = "Process Pictures (Crop)";
            this.butProcessPicturesCrop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butProcessPicturesCrop_Click);
            // 
            // PowerPointGraphics
            // 
            this.Name = "PowerPointGraphics";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabAppBuilderGraphics);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PowerPointGraphicsRibbon_Load);
            this.tabAppBuilderGraphics.ResumeLayout(false);
            this.tabAppBuilderGraphics.PerformLayout();
            this.grpPowerPointGraphicsTools.ResumeLayout(false);
            this.grpPowerPointGraphicsTools.PerformLayout();
            this.grpGraphicsSettings.ResumeLayout(false);
            this.grpGraphicsSettings.PerformLayout();
            this.grpSettings.ResumeLayout(false);
            this.grpSettings.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAppBuilderGraphics;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPowerPointGraphicsTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butLoadPictures;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBaseFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editFileExtension;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editGraphicsFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butLoadDefaults;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editDefaultsFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butBrowseGraphicsFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butSavePictures;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butProcessPicturesWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butSaveDefaults;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butSetDefaults;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpGraphicsSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPicturePositionWidthPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPicturePositionHeightPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPicturePositionLeftPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPicturePositionTopPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPictureCropWidthPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPictureCropHeightPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPictureCropLeftPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPictureCropTopPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkPositionPicture;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkCropPicture;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboGraphicsExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butClearSlides;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPictureOriginalWidthPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPictureOriginalHeightPix;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butProcessPicturesTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butProcessPicturesLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butProcessPicturesHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butBrowseBaseFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butProcessPicturesCrop;

    }

    partial class ThisRibbonCollection
    {
        internal PowerPointGraphics PowerPointGraphicsRibbon
        {
            get { return this.GetRibbon<PowerPointGraphics>(); }
        }
    }
}
