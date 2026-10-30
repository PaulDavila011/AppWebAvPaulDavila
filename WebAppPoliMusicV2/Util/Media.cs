using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web;

namespace WebAppPoliMusicV2.Util
{
    public class Media
    {
        /// <summary>
        /// Translate a path saved in the database ("~/songFiles/x.mp3") to the real
        /// path inside the media folder ("~/SongFilesAndImages_resources/songFiles/x.mp3")
        /// </summary>
        /// <param name="dbPath">Path as stored in the database</param>
        /// <returns>Application relative path of the file in the media folder</returns>
        public static string ToMediaPath(string dbPath)
        {
            if (dbPath == null || dbPath == string.Empty)
            {
                return string.Empty;
            }
            string mediaRoot = ConfigurationManager.AppSettings["mediaRootPath"];
            if (mediaRoot == null || mediaRoot == string.Empty)
            {
                return dbPath;
            }
            if (dbPath.StartsWith("~/") && !dbPath.StartsWith(mediaRoot, StringComparison.OrdinalIgnoreCase))
            {
                return mediaRoot + dbPath.Substring(2);
            }
            return dbPath;
        }
    }

}
