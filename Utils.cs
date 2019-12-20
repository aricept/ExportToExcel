using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.IO;
using System.Linq;
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

            if (Directory.Exists(path))
            {
                truePath = path;
                exists = true;
            } 
            else
            {
                try
                {
                    truePath = HostingEnvironment.MapPath(tempPath);
                    exists = true;
                }
                catch (ArgumentException e)
                {
                    
                }
            }

            return truePath;
        }

        /// <summary>
        /// Returns the <c>DisplayNameAttribute.DisplayName</c> or <c>DisplayAttribute.Name</c> for a Member. Used by <c>XlBlankSource</c> for creating header values.
        /// </summary>
        /// <param name="member">The Member to check for attribute</param>
        /// <returns></returns>
        public static string GetDisplayName(this MemberInfo member, Type type)
        {
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

            var metaTypeAttr = (MetadataTypeAttribute)type.GetCustomAttributes(typeof(MetadataTypeAttribute), true).OfType<MetadataTypeAttribute>().FirstOrDefault();
            if (metaTypeAttr != null)
            {
                var metaType = metaTypeAttr.MetadataClassType;
                var prop = metaType.GetMember(member.Name).FirstOrDefault();
                var metaAttr = prop.GetCustomAttribute<DisplayAttribute>();

                if (metaAttr != null)
                {
                    return metaAttr.Name;
                }
            }

            return string.Empty;
        }

        public static bool XlIgnore(this MemberInfo member, Type type)
        {
            var ignore = (XlIgnoreAttribute)Attribute.GetCustomAttribute(member, typeof(XlIgnoreAttribute));
            if (ignore != null)
            {
                return true;
            }

            var metaTypeAttr = (MetadataTypeAttribute)type.GetCustomAttributes(typeof(MetadataTypeAttribute), true).OfType<MetadataTypeAttribute>().FirstOrDefault();
            if(metaTypeAttr != null)
            {
                var metaType = metaTypeAttr.MetadataClassType;
                var prop = metaType.GetMember(member.Name).FirstOrDefault();
                var metaAttr = prop.GetCustomAttribute<XlIgnoreAttribute>();

                if (metaAttr != null)
                {
                    return true;
                }
            }

            return false;
        }

        public static DataType GetDataType(this MemberInfo member, Type type)
        {
            var dataType = (DataTypeAttribute)Attribute.GetCustomAttribute(member, typeof(DataTypeAttribute));
            if (dataType != null)
            {
                return dataType.DataType;
            }

            var metaTypeAttr = (MetadataTypeAttribute)type.GetCustomAttributes(typeof(MetadataTypeAttribute), true).OfType<MetadataTypeAttribute>().FirstOrDefault();
            if (metaTypeAttr != null)
            {
                var metaType = metaTypeAttr.MetadataClassType;
                var prop = metaType.GetMember(member.Name).FirstOrDefault();
                var metaAttr = prop.GetCustomAttribute<DataTypeAttribute>();

                if (metaAttr != null)
                {
                    return metaAttr.DataType;
                }
            }

            return DataType.Text;
        }
    }
}
