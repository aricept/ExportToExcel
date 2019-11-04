using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Web.Hosting;

namespace ExportToExcel
{
    public static class Utils
    {
        /// <summary>
        /// Determines if a provided string can be used to locate a valid directory. 
        /// </summary>
        /// <param name="path">String to check for path. May be an AppSettings key, virtual path, or absolute path.</param>
        /// <param name="exists">Outpout parameter, true if the directory is found, false if it was not.</param>
        /// <returns>The full path based on the string provided. May be null if not found.</returns>
        public static string GetDirPathIfExists(string path, out bool exists)
        {
            exists = false;
            var truePath = string.Empty;
            var tempPath = ConfigurationManager.AppSettings[path];

            if (tempPath == null)
            {
                tempPath = path;
            }

            tempPath = HostingEnvironment.MapPath(tempPath);

            if (Directory.Exists(tempPath))
            {
                truePath = tempPath;
                exists = true;
            }
            else if (Directory.Exists(path))
            {
                truePath = path;
                exists = true;
            }

            return truePath;
        }

        public static string GetDisplayName(this MemberInfo member)
        {
            string name;
            var displayName = (DisplayNameAttribute)Attribute.GetCustomAttribute(member, typeof(DisplayNameAttribute));
            if (displayName != null)
            {
                return displayName.DisplayName;
            }

            var displayAttrName = (DisplayAttribute)Attribute.GetCustomAttribute(member, typeof(DisplayAttribute));
            if (displayAttrName != null)
            {
                return displayAttrName.Name;
            }

            return string.Empty;
        }
    }
}
