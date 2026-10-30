using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    public class Song
    {
        private string strConnectionString;
        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }
        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }
        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }
        public Entities.Song ReadById(int id)
        {
            DataTable dtSong = new DAL.Song(strConnectionString).ReadById(id);
            if (dtSong.Rows.Count == 0)
            {
                return null;
            }
            DataRow row = dtSong.Rows[0];
            return new Entities.Song(
                Convert.ToInt32(row["ID_SONG"]),
                row["SONG_NAME"].ToString(),
                row["SONG_PATH"].ToString(),
                Convert.ToInt32(row["PLAYS"]));
        }

        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }

        public int Delete(int id)
        {
            return new DAL.Song(strConnectionString).Delete(id);
        }
        public int CheckExistSongName(string songName, int idSongExcluded)
        {
            return new DAL.Song(strConnectionString).CheckExistSongName(songName, idSongExcluded);
        }
        public string ValidationsDuplicated(Entities.Song song)
        {
            int returnValue = CheckExistSongName(song.SongName, song.ID);
            if (returnValue != 0)
            {
                return "A song with that name already exists!";
            }
            return string.Empty;
        }
        public string ValidationsDelete(int id)
        {
            int references = new DAL.Song(strConnectionString).CountReferences(id);
            if (references != 0)
            {
                return "The song cannot be deleted because it is used in albums, artists or playlists!";
            }
            return string.Empty;
        }
    }


}
