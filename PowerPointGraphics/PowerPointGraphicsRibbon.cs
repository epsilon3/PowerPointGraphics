using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Drawing;
using System.IO;

namespace PowerPointGraphics
{
    public partial class PowerPointGraphics
    {
        //  Initialization

        string mGraphicsExportMode = "Picture";
        bool mPositionPicture = true;
        bool mCropPicture = false;
        string mBaseFolder = "";
        string mDefaultsFileName = "";
        string mGraphicsFolder = "";
        string mFileExtension = "";
        int mPictureOriginalWidthPix = 0;
        int mPictureOriginalHeightPix = 0;
        int mPicturePositionLeftPix = 0;
        int mPicturePositionTopPix = 0;
        int mPicturePositionWidthPix = 0;
        int mPicturePositionHeightPix = 0;
        int mPictureCropWidthPix = 0;
        int mPictureCropHeightPix = 0;
        int mPictureCropLeftPix = 0;
        int mPictureCropTopPix = 0;

        List<string> mPictureFilePathNames = null;

        private void PowerPointGraphicsRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                mPictureFilePathNames = new List<string>();

                DoLoadDefaults();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [PowerPointGraphicsRibbon_Load]: " + ex);
            }

            return;
        }

        //  Interface

        private void butLoadPictures_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoLoadPictures();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [PowerPointGraphicsRibbon_Load]: " + ex);
            }
        }

        private void butSavePictures_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoSavePictures();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butSavePictures_Click]: " + ex);
            }
        }

        private void butProcessPicturesCrop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoProcessPictures("Crop");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butProcessPicturesCrop_Click]: " + ex);
            }
        }

        private void butProcessPicturesTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoProcessPictures("Top");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butProcessPicturesTop_Click]: " + ex);
            }
        }

        private void butProcessPicturesLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoProcessPictures("Left");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butProcessPicturesLeft_Click]: " + ex);
            }
        }

        private void butProcessPicturesWidth_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoProcessPictures("Width");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butProcessPicturesWidth_Click]: " + ex);
            }
        }

        private void butProcessPicturesHeight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoProcessPictures("Height");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butProcessPicturesHeight_Click]: " + ex);
            }
        }

        private void butClearSlides_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoClearSlides();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butClearSlides_Click]: " + ex);
            }
        }

        private void butLoadDefaults_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoLoadDefaults();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butLoadDefaults_Click]: " + ex);
            }
        }

        private void butSaveDefaults_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string strFileText = "";

                DoGetDefaults();

                strFileText = DoGetDefaultFileText();

                DoWriteFileText(mBaseFolder + "\\" + mDefaultsFileName, strFileText);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butSaveDefaults_Click]: " + ex);
            }
        }

        private void butSetDefaults_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DoSetDefaults();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butSetDefaults_Click]: " + ex);
            }
        }

        private void butBrowseBaseFolder_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FolderBrowserDialog fbdGet = null;
                DialogResult drGet;

                fbdGet = new FolderBrowserDialog();
                if (fbdGet != null)
                {
                    //fbdGet.RootFolder = System.Environment.SpecialFolder.MyComputer;

                    drGet = fbdGet.ShowDialog();
                    if (drGet.Equals(DialogResult.OK))
                    {
                        mBaseFolder = fbdGet.SelectedPath;
                        editBaseFolder.Text = mBaseFolder;
                    }
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [butBrowseBaseFolder_Click]: folder browser dialog is not set.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butBrowseBaseFolder_Click]: " + ex);
            }
        }

        private void butBrowseGraphicsFolder_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FolderBrowserDialog fbdGet = null;
                DialogResult drGet;
                
                fbdGet = new FolderBrowserDialog();
                if (fbdGet != null)
                {
                    //fbdGet.RootFolder = System.Environment.SpecialFolder.MyComputer;

                    drGet = fbdGet.ShowDialog();
                    if(drGet.Equals(DialogResult.OK))
                    {
                        mGraphicsFolder = fbdGet.SelectedPath;
                        editGraphicsFolder.Text = mBaseFolder;
                    }
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [butBrowseGraphicsFolder_Click]: folder browser dialog is not set.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [butBrowseGraphicsFolder_Click]: " + ex);
            }
        }

        private void chkPositionPicture_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (chkPositionPicture.Checked)
                    mPositionPicture = true;
                else
                    mPositionPicture = false;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [chkPositionPicture_Click]: " + ex);
            }
        }

        private void chkCropPicture_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (chkCropPicture.Checked)
                    mCropPicture = true;
                else
                    mCropPicture = false;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [chkCropPicture_Click]: " + ex);
            }
        }

        private void comboGraphicsExport_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (comboGraphicsExport.Text.Equals("Export Picture(s) to Graphic(s) (Excludes Slide Background)", StringComparison.OrdinalIgnoreCase))
                    mGraphicsExportMode = "Picture";
                else
                    mGraphicsExportMode = "Slide";
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [comboGraphicsExport_TextChanged]" + ex);
            }
        }
        
        private void editBaseFolder_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                mBaseFolder = editBaseFolder.Text;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editBaseFolder_TextChanged]" + ex);
            }
        }

        private void editDefaultsFileName_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                mDefaultsFileName = editDefaultsFileName.Text;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editDefaultsFileName_TextChanged]" + ex);
            }
        }

        private void editGraphicsFolder_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                mGraphicsFolder = editGraphicsFolder.Text;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editGraphicsFolder_TextChanged]" + ex);
            }
        }

        private void editFileExtension_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                mFileExtension = editFileExtension.Text;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editFileExtension_TextChanged]" + ex);
            }
        }
        
        private void editPictureOriginalWidthPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPictureOriginalWidthPix.Text.Equals(""))
                    mPictureOriginalWidthPix = Int32.Parse(editPictureOriginalWidthPix.Text);
                else
                {
                    mPictureOriginalWidthPix = 0;
                    editPictureOriginalWidthPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPictureOriginalWidthPix_TextChanged]" + ex);
            }
        }

        private void editPictureOriginalHeightPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPictureOriginalHeightPix.Text.Equals(""))
                    mPictureOriginalHeightPix = Int32.Parse(editPictureOriginalHeightPix.Text);
                else
                {
                    mPictureOriginalHeightPix = 0;
                    editPictureOriginalHeightPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPictureOriginalHeightPix_TextChanged]" + ex);
            }
        }

        private void editPicturePositionWidthPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPicturePositionWidthPix.Text.Equals(""))
                    mPicturePositionWidthPix = Int32.Parse(editPicturePositionWidthPix.Text);
                else
                {
                    mPicturePositionWidthPix = 0;
                    editPicturePositionWidthPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPicturePositionWidthPix_TextChanged]" + ex);
            }
        }

        private void editPicturePositionHeightPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPicturePositionHeightPix.Text.Equals(""))
                    mPicturePositionHeightPix = Int32.Parse(editPicturePositionHeightPix.Text);
                else
                {
                    mPicturePositionHeightPix = 0;
                    editPicturePositionHeightPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPicturePositionHeightPix_TextChanged]" + ex);
            }
        }

        private void editPicturePositionLeftPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPicturePositionLeftPix.Text.Equals(""))
                    mPicturePositionLeftPix = Int32.Parse(editPicturePositionLeftPix.Text);
                else
                {
                    mPicturePositionLeftPix = 0;
                    editPicturePositionLeftPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPicturePositionLeftPix_TextChanged]" + ex);
            }
        }

        private void editPicturePositionTopPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPicturePositionTopPix.Text.Equals(""))
                    mPicturePositionTopPix = Int32.Parse(editPicturePositionTopPix.Text);
                else
                {
                    mPicturePositionTopPix = 0;
                    editPicturePositionTopPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPicturePositionTopPix_TextChanged]" + ex);
            }
        }

        private void editPictureCropLeftPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPictureCropLeftPix.Text.Equals(""))
                    mPictureCropLeftPix = Int32.Parse(editPictureCropLeftPix.Text);
                else
                {
                    mPictureCropLeftPix = 0;
                    editPictureCropLeftPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPictureCropLeftPix_TextChanged]" + ex);
            }
        }

        private void editPictureCropTopPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPictureCropTopPix.Text.Equals(""))
                    mPictureCropTopPix = Int32.Parse(editPictureCropTopPix.Text);
                else
                {
                    mPictureCropTopPix = 0;
                    editPictureCropTopPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPictureCropTopPix_TextChanged]" + ex);
            }
        }

        private void editPictureCropWidthPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPictureCropWidthPix.Text.Equals(""))
                    mPictureCropWidthPix = Int32.Parse(editPictureCropWidthPix.Text);
                else
                {
                    mPictureCropWidthPix = 0;
                    editPictureCropWidthPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPictureCropWidthPix_TextChanged]" + ex);
            }
        }

        private void editPictureCropHeightPix_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!editPictureCropHeightPix.Text.Equals(""))
                    mPictureCropHeightPix = Int32.Parse(editPictureCropHeightPix.Text);
                else
                {
                    mPictureCropHeightPix = 0;
                    editPictureCropHeightPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [editPictureCropHeightPix_TextChanged]" + ex);
            }
        }

        //  Implementation

        private float DoGetPPTDims(float fIn)
        {
            try
            {
                return (float)(fIn / 1.3333333333333);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetPPTDims]" + ex);
            }

            return 0;
        }

        private float DoGetInchesFromPix(float fIn)
        {
            try
            {
                return (float)(fIn / 96);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetInchesFromPix]" + ex);
            }

            return 0;
        }

        private float DoGetPixFromInches(float fIn)
        {
            try
            {
                return (float)(fIn * 96);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetPixFromInches]" + ex);
            }

            return 0;
        }

        private float DoGetAbsoluteValue(float fIn)
        {
            try
            {
                if (fIn < 0)
                    return fIn * -1;
                else
                    return fIn;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetAbsoluteValue]" + ex);
            }

            return 0;
        }

        private void DoLoadDefaults()
        {
            string strDefaults = "";
            string[] arLines = null;
            string[] arKeyValue = null;

            try
            {
                mBaseFolder = editBaseFolder.Text;
                mDefaultsFileName = editDefaultsFileName.Text;

                strDefaults = DoGetFileText(mBaseFolder + "\\" + mDefaultsFileName);
                if (strDefaults != "")
                {
                    strDefaults = strDefaults.Replace("\r", "");
                    arLines = strDefaults.Split('\n');
                    if (arLines != null)
                    {
                        for (int nCount = 0; nCount < arLines.Count(); nCount++)
                        {
                            if (!arLines[nCount].Equals(""))
                            {
                                arKeyValue = arLines[nCount].Split('|');
                                if (arKeyValue != null)
                                {
                                    if (arKeyValue.Count() >= 2)
                                        DoSetDefault(arKeyValue[0], arKeyValue[1]);
                                    else if (arKeyValue.Count() >= 1)
                                        DoSetDefault(arKeyValue[0], "");
                                }
                            }
                        }
                    }
                }

                DoSetDefaults();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadDefaults]: " + ex);
            }

        }

        private void DoSetDefault(string strKey, string strValue)
        {
            try
            {
                if (strKey == "Graphics_Export_Mode") mGraphicsExportMode = strValue;
                else if (strKey == "Position_Picture")
                {
                    if (strValue.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                        mPositionPicture = true;
                    else
                        mPositionPicture = false;
                }
                else if (strKey == "Picture_Crop_Picture")
                {
                    if (strValue.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                        mCropPicture = true;
                    else
                        mCropPicture = false;
                }
                else if (strKey == "Base_Folder") mBaseFolder = strValue;
                else if (strKey == "Defaults_File_Name") mDefaultsFileName = strValue;
                else if (strKey == "Graphics_Folder") mGraphicsFolder = strValue;
                else if (strKey == "File_Extension") mFileExtension = strValue;
                else if (strKey == "Picture_Original_Width_Pix")
                {
                    if (!strValue.Equals(""))
                        mPictureOriginalWidthPix = Int32.Parse(strValue);
                    else
                        mPictureOriginalWidthPix = 0;
                }
                else if (strKey == "Picture_Original_Height_Pix")
                {
                    if (!strValue.Equals(""))
                        mPictureOriginalHeightPix = Int32.Parse(strValue);
                    else
                        mPictureOriginalHeightPix = 0;
                }
                else if (strKey == "Picture_Position_Width_Pix")
                {
                    if (!strValue.Equals(""))
                        mPicturePositionWidthPix = Int32.Parse(strValue);
                    else
                        mPicturePositionWidthPix = 0;
                }
                else if (strKey == "Picture_Position_Height_Pix")
                {
                    if (!strValue.Equals(""))
                        mPicturePositionHeightPix = Int32.Parse(strValue);
                    else
                        mPicturePositionHeightPix = 0;
                }
                else if (strKey == "Picture_Position_Left_Pix")
                {
                    if (!strValue.Equals(""))
                        mPicturePositionLeftPix = Int32.Parse(strValue);
                    else
                        mPicturePositionLeftPix = 0;
                }
                else if (strKey == "Picture_Position_Top_Pix")
                {
                    if (!strValue.Equals(""))
                        mPicturePositionTopPix = Int32.Parse(strValue);
                    else
                        mPicturePositionTopPix = 0;
                }
                else if (strKey == "Picture_Crop_Width_Pix")
                {
                    if (!strValue.Equals(""))
                        mPictureCropWidthPix = Int32.Parse(strValue);
                    else
                        mPictureCropWidthPix = 0;
                }
                else if (strKey == "Picture_Crop_Height_Pix")
                {
                    if (!strValue.Equals(""))
                        mPictureCropHeightPix = Int32.Parse(strValue);
                    else
                        mPictureCropHeightPix = 0;
                }
                else if (strKey == "Picture_Crop_Left_Pix")
                {
                    if (!strValue.Equals(""))
                        mPictureCropLeftPix = Int32.Parse(strValue);
                    else
                        mPictureCropLeftPix = 0;
                }
                else if (strKey == "Picture_Crop_Top_Pix")
                {
                    if (!strValue.Equals(""))
                        mPictureCropTopPix = Int32.Parse(strValue);
                    else
                        mPictureCropTopPix = 0;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSetDefault]: " + ex);
            }

            return;
        }

        private void DoSetDefaults()
        {
            try
            {
                if (mGraphicsExportMode.Equals("Picture", StringComparison.OrdinalIgnoreCase))
                    comboGraphicsExport.Text = "Export Picture(s) to Graphic(s) (Excludes Slide Background)";
                else
                    comboGraphicsExport.Text = "Export Slide(s) to Graphic(s) (Includes Slide Background)";
                chkPositionPicture.Checked = mPositionPicture;
                chkCropPicture.Checked = mCropPicture;
                //editBaseFolder.Text = mBaseFolder;
                //editDefaultsFileName.Text = mDefaultsFileName;
                editGraphicsFolder.Text = mGraphicsFolder;
                editFileExtension.Text = mFileExtension;
                editPictureOriginalWidthPix.Text = mPictureOriginalWidthPix.ToString();
                editPictureOriginalHeightPix.Text = mPictureOriginalHeightPix.ToString();
                editPicturePositionWidthPix.Text = mPicturePositionWidthPix.ToString();
                editPicturePositionHeightPix.Text = mPicturePositionHeightPix.ToString();
                editPicturePositionLeftPix.Text = mPicturePositionLeftPix.ToString();
                editPicturePositionTopPix.Text = mPicturePositionTopPix.ToString();
                editPictureCropWidthPix.Text = mPictureCropWidthPix.ToString();
                editPictureCropHeightPix.Text = mPictureCropHeightPix.ToString();
                editPictureCropLeftPix.Text = mPictureCropLeftPix.ToString();
                editPictureCropTopPix.Text = mPictureCropTopPix.ToString();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSetDefaults]: " + ex);
            }
        }

        private void DoGetDefaults()
        {
            try
            {
                if (comboGraphicsExport.Text.Equals("Export Picture(s) to Graphic(s) (Excludes Slide Background)"))
                    mGraphicsExportMode = "Picture";
                else
                    mGraphicsExportMode = "Slide";

                mPositionPicture = chkPositionPicture.Checked;
                mCropPicture = chkCropPicture.Checked;

                mBaseFolder = editBaseFolder.Text;
                mDefaultsFileName = editDefaultsFileName.Text;
                mGraphicsFolder = editGraphicsFolder.Text;
                mFileExtension = editFileExtension.Text;

                if (!editPictureOriginalWidthPix.Text.Equals(""))
                    mPictureOriginalWidthPix = Int32.Parse(editPictureOriginalWidthPix.Text);
                else
                {
                    mPictureOriginalWidthPix = 0;
                    editPictureOriginalWidthPix.Text = "0";
                }

                if(!editPictureOriginalHeightPix.Text.Equals(""))
                    mPictureOriginalHeightPix = Int32.Parse(editPictureOriginalHeightPix.Text);
                else
                {
                    mPictureOriginalHeightPix = 0;
                    editPictureOriginalHeightPix.Text = "0";
                }

                if(!editPicturePositionWidthPix.Text.Equals(""))
                    mPicturePositionWidthPix = Int32.Parse(editPicturePositionWidthPix.Text);
                else
                {
                    mPicturePositionWidthPix = 0;
                    editPicturePositionWidthPix.Text = "0";
                }

                if(!editPicturePositionHeightPix.Text.Equals(""))
                    mPicturePositionHeightPix = Int32.Parse(editPicturePositionHeightPix.Text);
                else
                {
                    mPicturePositionHeightPix = 0;
                    editPicturePositionHeightPix.Text = "0";
                }

                if(!editPicturePositionLeftPix.Text.Equals(""))
                    mPicturePositionLeftPix = Int32.Parse(editPicturePositionLeftPix.Text);
                else
                {
                    mPicturePositionLeftPix = 0;
                    editPicturePositionLeftPix.Text = "0";
                }

                if(!editPicturePositionTopPix.Text.Equals(""))
                    mPicturePositionTopPix = Int32.Parse(editPicturePositionTopPix.Text);
                else
                {
                    mPicturePositionTopPix = 0;
                    editPicturePositionTopPix.Text = "0";
                }

                if(!editPictureCropWidthPix.Text.Equals(""))
                    mPictureCropWidthPix = Int32.Parse(editPictureCropWidthPix.Text);
                else
                {
                    mPictureCropWidthPix = 0;
                    editPictureCropWidthPix.Text = "0";
                }

                if(!editPictureCropHeightPix.Text.Equals(""))
                    mPictureCropHeightPix = Int32.Parse(editPictureCropHeightPix.Text);
                else
                {
                    mPictureCropHeightPix = 0;
                    editPictureCropHeightPix.Text = "0";
                }

                if(!editPictureCropLeftPix.Text.Equals(""))
                    mPictureCropLeftPix = Int32.Parse(editPictureCropLeftPix.Text);
                else
                {
                    mPictureCropLeftPix = 0;
                    editPictureCropLeftPix.Text = "0";
                }

                if(!editPictureCropTopPix.Text.Equals(""))
                    mPictureCropTopPix = Int32.Parse(editPictureCropTopPix.Text);
                else
                {
                    mPictureCropTopPix = 0;
                    editPictureCropTopPix.Text = "0";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetDefaults]: " + ex);
            }
        }

        private string DoGetDefaultFileText()
        {
            string strReturn = "";

            try
            {
                strReturn = "Graphics_Export_Mode|" + mGraphicsExportMode + "\r\n";
                strReturn += "Position_Picture|";
                if (mPositionPicture)
                    strReturn += "Yes\r\n";
                else
                    strReturn += "No\r\n";
                strReturn += "Picture_Crop_Picture|";
                if (mCropPicture)
                    strReturn += "Yes\r\n";
                else
                    strReturn += "No\r\n";
                strReturn += "Base_Folder|" + mBaseFolder + "\r\n";
                strReturn += "Defaults_File_Name|" + mDefaultsFileName + "\r\n";
                strReturn += "Graphics_Folder|" + mGraphicsFolder + "\r\n";
                strReturn += "File_Extension|" + mFileExtension + "\r\n";
                strReturn += "Picture_Original_Width_Pix|" + mPictureOriginalWidthPix.ToString() + "\r\n";
                strReturn += "Picture_Original_Height_Pix|" + mPictureOriginalHeightPix.ToString() + "\r\n";
                strReturn += "Picture_Position_Width_Pix|" + mPicturePositionWidthPix.ToString() + "\r\n";
                strReturn += "Picture_Position_Height_Pix|" + mPicturePositionHeightPix.ToString() + "\r\n";
                strReturn += "Picture_Position_Left_Pix|" + mPicturePositionLeftPix.ToString() + "\r\n";
                strReturn += "Picture_Position_Top_Pix|" + mPicturePositionTopPix.ToString() + "\r\n";
                strReturn += "Picture_Crop_Width_Pix|" + mPictureCropWidthPix.ToString() + "\r\n";
                strReturn += "Picture_Crop_Height_Pix|" + mPictureCropHeightPix.ToString() + "\r\n";
                strReturn += "Picture_Crop_Left_Pix|" + mPictureCropLeftPix.ToString() + "\r\n";
                strReturn += "Picture_Crop_Top_Pix|" + mPictureCropTopPix.ToString() + "\r\n";
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetDefaultFileText]: " + ex);
            }

            return strReturn;
        }

        private bool DoWriteFileText(string strFilePathName, string strFileText)
        {
            try
            {
                if (!strFilePathName.Equals(""))
                {
                    StreamWriter strmWriter = new StreamWriter(strFilePathName);
                    try
                    {
                        strmWriter.Write(strFileText);
                    }
                    finally
                    {
                        strmWriter.Close();
                    }

                    return true;
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]:  File path name is empty.");
            }
            catch (OutOfMemoryException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (ArgumentNullException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (ArgumentException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (FileNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (DirectoryNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (IOException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoWriteFileText]: {" + strFilePathName + "} " + ex);
            }

            return false;
        }

        private string DoGetFileText(string strFilePathName)
        {
            string strFileText = "";

            try
            {
                StreamReader strmReader = new StreamReader(strFilePathName);
                try
                {
                    strFileText = strmReader.ReadToEnd();
                }
                finally
                {
                    strmReader.Close();
                }
            }
            catch (OutOfMemoryException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (ArgumentNullException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (ArgumentException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (FileNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (DirectoryNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (IOException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoGetFileText]: {" + strFilePathName + "} " + ex);
            }

            return strFileText;
        }

        private PowerPoint.Shape DoLoadShape(string strShapeType, string strFilePathName, string strName, float fLeft, float fTop, float fWidth, float fHeight)
        {
            PowerPoint.Shape shapeReturn = null;

            try
            {
                // get last slide
                PowerPoint.Slide sldLoad = Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActivePresentation.Slides.Count];

                if (strShapeType == "Rectangle")
                {
                    shapeReturn = sldLoad.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, fLeft, fTop, fWidth, fHeight);
                }
                else if (strShapeType == "Line")
                {
                    shapeReturn = sldLoad.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeLineCallout1NoBorder, fLeft, fTop, fWidth, fHeight);
                }
                else if (strShapeType == "TextBox")
                {
                    shapeReturn = sldLoad.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, fLeft, fTop, fWidth, fHeight);
                }
                else if (strShapeType == "Picture")
                {
                    shapeReturn = sldLoad.Shapes.AddPicture(strFilePathName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, fLeft, fTop, fWidth, fHeight);//fWidth, fHeight);
                    //shapeReturn.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                }

                shapeReturn.Name = strName;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadShape] {" + strShapeType + " : " + strFilePathName + " : " + strName + "} failed to load shape: " + ex);
            }

            return shapeReturn;
        }

        private void DoLoadPictures()
        {
            try
            {
                string strFolder = "";
                string strName = "";
                float fLeft = 0;
                float fTop = 0;
                float fWidth = 0;
                float fHeight = 0;
                //int nConvert = 0;
                string[] arFilePath = null;

                if (mPictureFilePathNames != null)
                    mPictureFilePathNames.Clear();
                else
                    mPictureFilePathNames = new List<string>();

                DoGetDefaults();

                if(mPictureFilePathNames != null)
                { 
                    if (!mBaseFolder.Equals(""))
                    {
                        if (!mGraphicsFolder.Equals(""))
                            strFolder = mBaseFolder + "\\" + mGraphicsFolder;
                        else
                            strFolder = mBaseFolder;

                        if(mFileExtension.Equals(""))
                        {
                            mFileExtension = "*";
                            editFileExtension.Text = "*";
                        }

                        if (!mFileExtension.Equals(""))
                        {
                            foreach (string strFilePathName in Directory.GetFiles(mBaseFolder + "\\" + mGraphicsFolder, "*." + mFileExtension, SearchOption.AllDirectories))
                            {
                                arFilePath = strFilePathName.Split('\\');
                                if ((arFilePath != null) && (arFilePath.Count() > 0))
                                {
                                    strName = arFilePath[arFilePath.Count() - 1];
                                    if (!strName.Equals(""))
                                    {
                                        if(mPositionPicture)
                                        {
                                            fLeft = DoGetPPTDims((mPicturePositionWidthPix - mPictureOriginalWidthPix) / 2);
                                            fTop = DoGetPPTDims((mPicturePositionHeightPix - mPictureOriginalHeightPix) / 2);
                                        }
                                        else
                                        {
                                            fLeft = 0;
                                            fTop = 0;
                                        }
                                        fWidth = DoGetPPTDims(mPictureOriginalWidthPix);
                                        fHeight = DoGetPPTDims(mPictureOriginalHeightPix);

                                        PowerPoint.Shape shapeLoad = null;

                                        Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides.Count + 1, Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.CustomLayouts[1]);

                                        DoClearSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides.Count);

                                        shapeLoad = DoLoadShape("Picture", strFilePathName, "Picture", fLeft, fTop, fWidth, fHeight);

                                        //nConvert = (int)fLeft;
                                        //editPicturePositionLeftPix.Text = nConvert.ToString();
                                        //nConvert = (int)fTop;
                                        //editPicturePositionTopPix.Text = nConvert.ToString();
                                        //nConvert = (int)fWidth;
                                        //editPicturePositionWidthPix.Text = nConvert.ToString();
                                        //nConvert = (int)fHeight;
                                        //editPicturePositionHeightPix.Text = nConvert.ToString();

                                        mPictureFilePathNames.Add(strFilePathName);
                                    }
                                    else
                                        System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: file name is empty.");
                                }
                                else
                                    System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: file path name is empty.");
                            }
                        }
                        else
                            System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: graphics folder path is empty.");
                    }
                    else
                        System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: base folder path is empty.");
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: file path names list is not set.");
            }
            catch (OutOfMemoryException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
            catch (ArgumentNullException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
            catch (ArgumentException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
            catch (FileNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
            catch (DirectoryNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
            catch (IOException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoLoadPictures]: " + ex);
            }
        }

        private void DoProcessPictures(string strMode)
        {
            try
            {
                DoGetDefaults();

                if (strMode.Equals("Crop", StringComparison.OrdinalIgnoreCase))
                {
                    for (int nCount = 1; nCount <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count; nCount++)
                    {
                        PowerPoint.Shape shapeProcess = null;

                        shapeProcess = Globals.ThisAddIn.Application.ActivePresentation.Slides[nCount].Shapes["Picture"];
                        if (shapeProcess != null)
                        {
                            shapeProcess.PictureFormat.Crop.ShapeLeft = DoGetPPTDims(mPictureCropLeftPix);
                            shapeProcess.PictureFormat.Crop.ShapeTop = DoGetPPTDims(mPictureCropTopPix);
                            shapeProcess.PictureFormat.Crop.ShapeWidth = DoGetPPTDims(mPictureCropWidthPix);
                            shapeProcess.PictureFormat.Crop.ShapeHeight = DoGetPPTDims(mPictureCropHeightPix);
                        }
                    }
                }
                else if(strMode.Equals("Left", StringComparison.OrdinalIgnoreCase))
                {
                    for(int nCount = 1 ; nCount <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count ; nCount++)
                    {
                        PowerPoint.Shape shapeProcess = null;

                        shapeProcess = Globals.ThisAddIn.Application.ActivePresentation.Slides[nCount].Shapes["Picture"];
                        if(shapeProcess != null)
                            shapeProcess.Left = DoGetPPTDims(mPicturePositionLeftPix);
                    }
                }
                else if (strMode.Equals("Top", StringComparison.OrdinalIgnoreCase))
                {
                    for (int nCount = 1; nCount <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count; nCount++)
                    {
                        PowerPoint.Shape shapeProcess = null;

                        shapeProcess = Globals.ThisAddIn.Application.ActivePresentation.Slides[nCount].Shapes["Picture"];
                        if (shapeProcess != null)
                            shapeProcess.Top = DoGetPPTDims(mPicturePositionTopPix);
                    }
                }
                else if (strMode.Equals("Width", StringComparison.OrdinalIgnoreCase))
                {
                    for (int nCount = 1; nCount <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count; nCount++)
                    {
                        PowerPoint.Shape shapeProcess = null;

                        shapeProcess = Globals.ThisAddIn.Application.ActivePresentation.Slides[nCount].Shapes["Picture"];
                        if (shapeProcess != null)
                            shapeProcess.Width = DoGetPPTDims(mPicturePositionWidthPix);
                    }
                }
                else if (strMode.Equals("Height", StringComparison.OrdinalIgnoreCase))
                {
                    for (int nCount = 1; nCount <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count; nCount++)
                    {
                        PowerPoint.Shape shapeProcess = null;

                        shapeProcess = Globals.ThisAddIn.Application.ActivePresentation.Slides[nCount].Shapes["Picture"];
                        if (shapeProcess != null)
                            shapeProcess.Height = DoGetPPTDims(mPicturePositionHeightPix);
                    }
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoProcessPictures]: {" + strMode + "} mode is not correctly set - it should be one of [Crop, Left, Top, Width, Height].");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoProcessPictures]: {" + strMode + "} " + ex);
            }
        }

        private void DoSavePictures()
        {
            try
            {
                if (mPictureFilePathNames != null)
                {
                    for (int nCount = 0; nCount < mPictureFilePathNames.Count; nCount++)
                    {
                        DoPictureExport(nCount+1, "Picture", mPictureFilePathNames.ElementAt(nCount));
                    }
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: file path names list is not set.");
            }
            catch (OutOfMemoryException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
            catch (ArgumentNullException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
            catch (ArgumentException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
            catch (FileNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
            catch (DirectoryNotFoundException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
            catch (IOException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSavePictures]: " + ex);
            }
        }

        private void DoClearSlides()
        {
            try
            {
                if (mPictureFilePathNames != null)
                    mPictureFilePathNames.Clear();
                else
                    mPictureFilePathNames = new List<string>();

                for (int nCount = Globals.ThisAddIn.Application.ActivePresentation.Slides.Count; nCount >= 1; nCount--)
                    Globals.ThisAddIn.Application.ActivePresentation.Slides[nCount].Delete();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoClearSlides]: " + ex);
            }
        }

        private void DoSwitchExport(int nIndex, string strShapeName, string strFilePathName, float fWidth, float fHeight)
        {
            try
            {
                if ((nIndex >= 1) && (nIndex <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count))
                {
                    // Export
                    if ((fWidth < 96) || (fHeight < 96))
                        DoPictureExport(nIndex, strShapeName, strFilePathName);
                    else
                        DoSlideExport(nIndex, strFilePathName);
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoSwitchExport]: slide index [" + nIndex.ToString() + "] is out of range [1:" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Count.ToString() + "], inclusive.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSwitchExport]: {" + strFilePathName + "} " + ex);

            }
        }

        private void DoPictureExport(int nIndex, string strShapeName, string strFilePathName)
        {
            try
            {
                if ((nIndex >= 1) && (nIndex <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count))
                {
                    // get first slide
                    PowerPoint.Slide sldMain = Globals.ThisAddIn.Application.ActivePresentation.Slides[nIndex];

                    // Export Image
                    sldMain.Shapes[strShapeName].Export(strFilePathName, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoSlideExport]: slide index [" + nIndex.ToString() + "] is out of range [1:" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Count.ToString() + "], inclusive.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoPictureExport]: {" + strFilePathName + "} " + ex);

            }
        }

        private void DoSlideExport(int nIndex, string strFilePathName)
        {
            try
            {
                if ((nIndex >= 1) && (nIndex <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count))
                {
                    // Get Slide
                    PowerPoint.Slide sldMain = Globals.ThisAddIn.Application.ActivePresentation.Slides[nIndex];

                    // Export
                    sldMain.Export(strFilePathName, strFilePathName.Substring(strFilePathName.Length - 3, 3));
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoSlideExport]: slide index [" + nIndex.ToString() + "] is out of range [1:" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Count.ToString() + "], inclusive.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoSlideExport]: {" + strFilePathName + "} " + ex);

            }
        }

        private void DoClearSlide(int nIndex)
        {
            try
            {
                if ((nIndex >= 1) && (nIndex <= Globals.ThisAddIn.Application.ActivePresentation.Slides.Count))
                {
                    PowerPoint.Slide sldMain = Globals.ThisAddIn.Application.ActivePresentation.Slides[nIndex];

                    for (int nCount = sldMain.Shapes.Count; nCount > 0; nCount--)
                        sldMain.Shapes[nCount].Delete();

                    GC.Collect();
                }
                else
                    System.Windows.Forms.MessageBox.Show("Error [DoClearSlide]: slide index [" + nIndex.ToString() + "] is out of range [1:" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Count.ToString() + "], inclusive.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error [DoClearSlide]: " + ex);
            }
        }
    }
}
