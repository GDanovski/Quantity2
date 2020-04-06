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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
namespace Quantity2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(@"Quantity2 - software for data analysis
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
 along with this program.If not, see<http://www.gnu.org/licenses/>.
-------------------------------------------------------------------");

            Console.WriteLine("Hello!");
            Console.WriteLine("");

            while (true)
            {
                int MaxColumnLength = 331;

                Console.WriteLine("ImageN:");
                if (!int.TryParse(Console.ReadLine(), out MaxColumnLength)) MaxColumnLength = 331;

                string dir, suf;
                //Get color
                Console.WriteLine("Suffix:");
                suf = Console.ReadLine();

                //Add work directory
                do
                {
                    Console.WriteLine("Work directory:");
                    dir = Console.ReadLine();
                    Console.Write("\n");
                }
                while (!Directory.Exists(dir));

                if (!Directory.Exists(dir))
                {
                    Console.WriteLine("Error dir!");
                    continue;
                }
                int startG = GetValue("Normalize - start G:");
                int stopG = GetValue("Normalize - stop G:");
                int startR = GetValue("Normalize - start R:");
                int stopR = GetValue("Normalize - stop R:");

                ExcelCode.CreateWorkbook(suf, MaxColumnLength, dir,startG,stopG,startR,stopR);
                
                Console.ReadKey();
            }
        }
        private static int GetValue(string name)
        {
            int result = 0;
            string str = "";
            do
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(" >>> " + name + ":");
                Console.ForegroundColor = ConsoleColor.Green;
                str = Console.ReadLine();

                if (!int.TryParse(str, out result))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(" >>> Incorrect Value!");
                }

                Console.ForegroundColor = ConsoleColor.White;
            }
            while (!int.TryParse(str, out result));

            Console.ForegroundColor = ConsoleColor.White;

            return result;
        }
    }
}
