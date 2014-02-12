using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using Microsoft.Office.Interop.Outlook;

namespace Tabbles.OutlookAddIn
{
    public static class Utils
    {
        public const int MAPI_E_COLLISION = -2147219964;

        private static readonly Dictionary<Outlook.OlCategoryColor, string> OutlookColorsStr = new Dictionary<Outlook.OlCategoryColor, string>()
        {
            {Outlook.OlCategoryColor.olCategoryColorBlack, "4F4F4F"},
            {Outlook.OlCategoryColor.olCategoryColorBlue, "9DB7E8"},
            {Outlook.OlCategoryColor.olCategoryColorDarkBlue, "2858A5"},
            {Outlook.OlCategoryColor.olCategoryColorDarkGray, "6F6F6F"},
            {Outlook.OlCategoryColor.olCategoryColorDarkGreen, "3F8F2B"},
            {Outlook.OlCategoryColor.olCategoryColorDarkMaroon, "93446B"},
            {Outlook.OlCategoryColor.olCategoryColorDarkOlive, "778B45"},
            {Outlook.OlCategoryColor.olCategoryColorDarkOrange, "E2620D"},
            {Outlook.OlCategoryColor.olCategoryColorDarkPeach, "C79930"},
            {Outlook.OlCategoryColor.olCategoryColorDarkPurple, "5C3FA3"},
            {Outlook.OlCategoryColor.olCategoryColorDarkRed, "C11A25"},
            {Outlook.OlCategoryColor.olCategoryColorDarkSteel, "6B7994"},
            {Outlook.OlCategoryColor.olCategoryColorDarkTeal, "329B7A"},
            {Outlook.OlCategoryColor.olCategoryColorDarkYellow, "B9B300"},
            {Outlook.OlCategoryColor.olCategoryColorGray, "BFBFBF"},
            {Outlook.OlCategoryColor.olCategoryColorGreen, "78D168"},
            {Outlook.OlCategoryColor.olCategoryColorMaroon, "DAAEC2"},
            {Outlook.OlCategoryColor.olCategoryColorNone, "FFFFFF"},
            {Outlook.OlCategoryColor.olCategoryColorOlive, "C6D2B0"},
            {Outlook.OlCategoryColor.olCategoryColorOrange, "F9BA89"},
            {Outlook.OlCategoryColor.olCategoryColorPeach, "F7DD8F"},
            {Outlook.OlCategoryColor.olCategoryColorPurple, "B5A1E2"},
            {Outlook.OlCategoryColor.olCategoryColorRed, "E7A1A2"},
            {Outlook.OlCategoryColor.olCategoryColorSteel, "DAD9DC"},
            {Outlook.OlCategoryColor.olCategoryColorTeal, "9FDCC9"},
            {Outlook.OlCategoryColor.olCategoryColorYellow, "FCFA90"}
        };

        private const string OutlookColorStrFollowUp = "F6532F";

        private static readonly string[] CategorySeparator = new string[] { ", " };

        private const string RegKeyTabbles = @"SOFTWARE\Yellow Blue Soft\Tabbles";
        private const string RegValueTabblesInstallDir = "installation_dir";

        private static readonly Dictionary<Outlook.OlCategoryColor, System.Drawing.Color> OutlookColorsRgb
            = new Dictionary<OlCategoryColor, System.Drawing.Color>();

        static Utils()
        {
            foreach (var outlookColorStr in OutlookColorsStr)
            {
                System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml("#" + outlookColorStr.Value);
                OutlookColorsRgb.Add(outlookColorStr.Key, color);
            }
        }

        public static OutlookVersion ParseMajorVersion(Outlook.Application outlookApplication)
        {
            string majorVersionString = outlookApplication.Version.Split(new char[] { '.' })[0];
            switch (majorVersionString)
            {
                case "11":
                    return OutlookVersion.OUTLOOK_2003;
                case "12":
                    return OutlookVersion.OUTLOOK_2007;
                case "14":
                    return OutlookVersion.OUTLOOK_2010;
                default:
                    return OutlookVersion.UNKNOWN;
            }
        }

        public static string GetOutlookPrefix()
        {
            string path = GetOutlookPath();
            return "\"" + path + @""" /select outlook:";
        }

        private static string GetOutlookPath()
        {
            // Fetch the Outlook Class ID
            var key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\Outlook.Application\\CLSID");
            var objOutlookClassID = key.GetValue("");
            var outlookClassId = ((string)objOutlookClassID).Trim();
            key.Dispose();

            // Using the class ID from above pull up the path
            var key2 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\CLSID\" + outlookClassId + @"\LocalServer32");
            var outlookPath = ((string)key2.GetValue("")).Trim();
            key2.Dispose();

            return outlookPath;
        }

        public static string GetTabblesInstallDir()
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(RegKeyTabbles);
            if (key == null)
            {
                key = Registry.CurrentUser.OpenSubKey(RegKeyTabbles);
            }

            if (key != null)
            {
                return key.GetValue(RegValueTabblesInstallDir) as string;
            }

            return null;
        }

        public static Outlook.OlCategoryColor GetOutlookColorFromRgb(string rgb)
        {
            if (!string.IsNullOrEmpty(rgb) && (rgb = rgb.Trim()).Length != 9)
            {
                return OlCategoryColor.olCategoryColorNone;
            }

            string rStr = rgb.Substring(3, 2);
            string gStr = rgb.Substring(5, 2);
            string bStr = rgb.Substring(7, 2);

            int r;
            int g;
            int b;
            try
            {
                r = Int32.Parse(rStr, System.Globalization.NumberStyles.HexNumber);
                g = Int32.Parse(gStr, System.Globalization.NumberStyles.HexNumber);
                b = Int32.Parse(bStr, System.Globalization.NumberStyles.HexNumber);

                OlCategoryColor olColor = OlCategoryColor.olCategoryColorNone;

                //calculate nearest color with two algorithms
                double minDistance = double.MaxValue;
                foreach (var olColorRgb in OutlookColorsRgb)
                {
                    int rDiff = r - olColorRgb.Value.R;
                    int gDiff = g - olColorRgb.Value.G;
                    int bDiff = b - olColorRgb.Value.B;

                    //1. consider RGB as a three-dimensional space and get the nearest color
                    double curDistance = Math.Sqrt(rDiff * rDiff + gDiff * gDiff + bDiff * bDiff);
                    //2. count the total difference of color and choose the nearest one
                    curDistance += Math.Abs(rDiff) + Math.Abs(gDiff) + Math.Abs(bDiff);
                    if (curDistance < minDistance)
                    {
                        minDistance = curDistance;
                        olColor = olColorRgb.Key;
                    }
                }

                return olColor;
            }
            catch (System.Exception)
            {
                return OlCategoryColor.olCategoryColorNone;
            }
        }

        public static string GetRgbFromOutlookColor(Outlook.OlCategoryColor color)
        {
            //Tabbles needs # and alpha value in addition to RGB
            return "#FF" + OutlookColorsStr[color];
        }

        public static string GetRgbForFlagRequest(string flagRequest)
        {
            return "#FF" + OutlookColorStrFollowUp;
        }

        /// <summary>
        /// Returns array of categories of given mail item.
        /// </summary>
        /// <param name="mail"></param>
        /// <returns></returns>
        public static string[] GetCategories(MailItem mail)
        {
            if (mail != null && mail.Categories != null)
            {
                return mail.Categories.Split(CategorySeparator, StringSplitOptions.None);
            }

            return new string[0];
        }

        /// <summary>
        /// Removes given folder from its parent.
        /// </summary>
        /// <param name="folder"></param>
        public static void RemoveFolder(Folder folder)
        {
            try
            {
                if (folder != null)
                {
                    object parentObj = folder.Parent;
                    if (parentObj is Folder)
                    {
                        Folders subfolders = ((Folder)parentObj).Folders;
                        for (int i = 1, length = subfolders.Count; i <= length; i++) //folder index is 1-based
                        {
                            if (subfolders[i].Name == folder.Name)
                            {
                                subfolders.Remove(i);
                            }
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                //ignore this exception
            }
        }

        public static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                }
                catch (System.Exception)
                {
                    //do nothing
                }
            }
        }
    }

    public enum OutlookVersion
    {
        OUTLOOK_2003,
        OUTLOOK_2007,
        OUTLOOK_2010,
        UNKNOWN
    }
}
