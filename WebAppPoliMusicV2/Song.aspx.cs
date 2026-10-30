using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
        string strSongFolder = ConfigurationManager.AppSettings["songFolderPath"];

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                lblUser.Text = HttpUtility.HtmlEncode(User.Identity.Name);
                ClearForm();
                LoadSongs();
            }
        }

        private void LoadSongs()
        {
            try
            {
                DataTable dtSongs = new BLL.Song(strConnString).Read();
                gvSongs.DataSource = dtSongs;
                gvSongs.DataBind();
                litCount.Text = dtSongs.Rows.Count + " song(s)";
            }
            catch (Exception ex)
            {
                gvSongs.DataSource = null;
                gvSongs.DataBind();
                ShowMessage("Error loading songs: " + ex.Message, false);
            }
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            string messageValidations = string.Empty;
            int idSong = Convert.ToInt32(hfIdSong.Value);
            bool isNew = idSong == 0;

            Entities.Song song = new Entities.Song();
            song.ID = idSong;
            song.SongName = txtSongName.Text.Trim();
            song.Plays = 0;

            messageValidations = SongFieldsRequiredValidations(song, isNew);
            if (messageValidations != string.Empty)
            {
                ShowMessage(messageValidations, false);
                return;
            }
            try
            {
                BLL.Song bllSong = new BLL.Song(strConnString);
                messageValidations = bllSong.ValidationsDuplicated(song);
                if (messageValidations != string.Empty)
                {
                    ShowMessage(messageValidations, false);
                    return;
                }
                if (fupSong.HasFile)
                {
                    string newSongName = BuildFileName(song.SongName);
                    messageValidations = new FileManagement().SaveSongOnServer(fupSong, Util.Media.ToMediaPath(strSongFolder), newSongName);
                    if (messageValidations != string.Empty)
                    {
                        ShowMessage(messageValidations, false);
                        return;
                    }
                    song.SongPath = strSongFolder + newSongName + Path.GetExtension(fupSong.FileName).ToLower();
                }

                if (isNew)
                {
                    bllSong.Save(song);
                    ShowMessage("Song registered successfully!", true);
                }
                else
                {
                    Entities.Song oldSong = bllSong.ReadById(idSong);
                    bllSong.Update(song);
                    if (oldSong != null && !string.IsNullOrEmpty(song.SongPath))
                    {
                        DeleteSongFile(oldSong.SongPath);
                    }
                    ShowMessage("Song updated successfully!", true);
                }
                ClearForm();
                LoadSongs();
            }
            catch (Exception ex)
            {
                ShowMessage("Error saving the song: " + ex.Message, false);
            }
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            ClearForm();
            pnlMessage.Visible = false;
        }

        protected void gvSongs_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName != "EditSong" && e.CommandName != "DeleteSong")
            {
                return;
            }
            int idSong = Convert.ToInt32(e.CommandArgument);
            try
            {
                BLL.Song bllSong = new BLL.Song(strConnString);
                if (e.CommandName == "EditSong")
                {
                    Entities.Song song = bllSong.ReadById(idSong);
                    if (song == null)
                    {
                        ShowMessage("The song no longer exists!", false);
                        LoadSongs();
                        return;
                    }
                    hfIdSong.Value = song.ID.ToString();
                    txtSongName.Text = song.SongName;
                    litFormTitle.Text = "Edit song";
                    lblFileRequired.Visible = false;
                    lblCurrentFile.Text = "Current file: " + HttpUtility.HtmlEncode(song.SongPath) + " (leave empty to keep it)";
                    btnSave.Text = "Update";
                    pnlMessage.Visible = false;
                }
                else
                {
                    string messageValidations = bllSong.ValidationsDelete(idSong);
                    if (messageValidations != string.Empty)
                    {
                        ShowMessage(messageValidations, false);
                        return;
                    }
                    Entities.Song song = bllSong.ReadById(idSong);
                    bllSong.Delete(idSong);
                    if (song != null)
                    {
                        DeleteSongFile(song.SongPath);
                    }
                    if (hfIdSong.Value == idSong.ToString())
                    {
                        ClearForm();
                    }
                    ShowMessage("Song deleted successfully!", true);
                    LoadSongs();
                }
            }
            catch (Exception ex)
            {
                ShowMessage("Error: " + ex.Message, false);
            }
        }

        public string SongFieldsRequiredValidations(Entities.Song song, bool isNew)
        {
            if (song.SongName == string.Empty)
            {
                return "Song name is required!";
            }
            if (song.SongName.Length > 50)
            {
                return "Song name must have 50 characters or less!";
            }
            if (isNew && !fupSong.HasFile)
            {
                return "Select a song file first!!";
            }
            return string.Empty;
        }

        /// <summary>
        /// Build a safe and unique file name based on the song name
        /// </summary>
        private string BuildFileName(string songName)
        {
            string cleanName = Regex.Replace(songName, @"[^A-Za-z0-9_\-]", string.Empty);
            if (cleanName == string.Empty)
            {
                cleanName = "song";
            }
            return cleanName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        private void DeleteSongFile(string songPath)
        {
            try
            {
                if (string.IsNullOrEmpty(songPath))
                {
                    return;
                }
                string serverPath = Server.MapPath(Util.Media.ToMediaPath(songPath));
                if (File.Exists(serverPath))
                {
                    File.Delete(serverPath);
                }
            }
            catch (Exception)
            {
                // The record is already updated, a missing file is not critical
            }
        }

        private void ClearForm()
        {
            hfIdSong.Value = "0";
            txtSongName.Text = string.Empty;
            litFormTitle.Text = "New song";
            lblFileRequired.Visible = true;
            lblCurrentFile.Text = string.Empty;
            btnSave.Text = "Save";
        }

        private void ShowMessage(string message, bool success)
        {
            pnlMessage.Visible = true;
            pnlMessage.CssClass = success ? "pm-flash pm-flash-success" : "pm-flash pm-flash-error";
            lblMessage.Text = HttpUtility.HtmlEncode(message);
        }
    }


}
