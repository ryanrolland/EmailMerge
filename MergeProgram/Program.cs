using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EmailMerge;
using System.IO;
namespace wtf
{
    class Program
    {
        /// <summary>
        /// Runs the Email Merge program to either merge pst2 into pst1, or into a new merged pst file.
        /// </summary>
        /// <param name="args">Inputs in order:
        /// 1- pst1 filename
        /// 2- pst2 filename
        /// 3- [optional] merged filename
        /// 4- [optional] duplicates filename from pst1
        /// 5- [optional] duplicates filename from pst2
        /// 6- [optional] boolean save duplicates
        /// 7- [optional] comma delim string of folders to ignore</param>
        static void Main(string[] args)
        {
            try
            {

                Console.WriteLine("Loading and Merging PST Files");


                PSTFile pstFile1 = new PSTFile(@"C:\dev\EmailMerge\working\010116.pst", "010116.pst");
                PSTFile pstFile2 = new PSTFile(@"C:\dev\EmailMerge\working\110915.pst", "110915.pst");

                PSTFile.MergePSTFiles(pstFile1,pstFile2, @"C:\dev\EmailMerge\working\merge1.pst", null,null,false,null);

                //PSTFile mergedPSTFile = new PSTFile(@"C:\dev\EmailMerge\working\merge1.pst", "Merged");





                //}

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
            Console.WriteLine("Merge Finished... press any key to exit");
            Console.ReadKey(true);
        }

       
    }
}
