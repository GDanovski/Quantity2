/*Quantity2 - software for data analysis
 Copyright(C) 2018  Georgi Danovski

 This program is free software: you can redistribute it and/or modify
 it under the terms of the GNU General Public License as published by
 the Free Software Foundation, either version 3 of the License, or
 (at your option) any later version.

 This program is distributed in the hope that it will be useful,
 but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
 GNU General Public License for more details.

 You should have received a copy of the GNU General Public License
 along with this program.If not, see<http://www.gnu.org/licenses/>.*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quantity2
{
    class FileReader
    {
        public static List<string> DigForFiles(string dir, string suf)
        {
           

            List<string> result = new List<string>();
            List<string> dirs = new List<string>() { dir };
            List<string> temp;
            suf += "_Results.txt";

            while (dirs.Count > 0)
            {
                temp = new List<string>();

                foreach (string str in dirs)
                {
                    foreach (string name in GetFiles(str,suf))
                        result.Add(name);

                    foreach (string name in GetDirectories(str))
                        temp.Add(name);
                }

                dirs = temp;
            }

            dirs = null;
            temp = null;

            return result;
        }
        private static List<string> GetFiles(string dir, string suf)
        {
            List<string> result = new List<string>();

            DirectoryInfo di = new DirectoryInfo(dir);

            foreach (var fi in di.GetFiles())
                if (fi.Extension == ".txt" && fi.FullName.EndsWith(suf))
                    result.Add(fi.FullName);

            di = null;

            return result;
        }
        private static List<string> GetDirectories(string dir)
        {
            List<string> result = new List<string>();

            DirectoryInfo di = new DirectoryInfo(dir);

            foreach (var fi in di.GetDirectories())
                result.Add(fi.FullName);

            di = null;

            return result;
        }
    }
}

