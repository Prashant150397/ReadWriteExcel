using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadWriteExcel
{
     class checkExcel
    {
        static void Main(string[] args)
        {


            int[] arr = { 5,23,19,4, 5, 6, 7 };

            int max = arr[0];
            int min = arr[0];

            for (int i = 1; i < arr.Length; i++)
            { 
                if (arr[i] > max)
                {
                    max = arr[i];
                }else if (arr[i] < min) { min= arr[i]; }
            }
            Console.WriteLine(max+ " max value");
            Console.WriteLine(min + " min value");
        }
        
        }
}
