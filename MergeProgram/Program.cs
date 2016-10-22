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
        /// Runs the Email Merge program to merge all pst files in the directory to search into the merge filename (pst). A new pst may be specified.
        /// It is possible you will be prompted to create the pst by specifying a directory for it to live in. The prompt is confusing because it acts like the "merge.pst" cannot be found if it
        /// doesn't exist.
        /// 
        /// I forked this project to solve an issue where I had about 20 different pst files that I wanted to consolidate into one. It looked like the general solutions out there all required you
        /// to pay for a small GUI utility with the exception of this project but the code needed to be modified slightly... hence this commit.
        /// 
        /// A had a couple of different issues:
        /// 1) The unique key that was being generated was throwing an exception on a couple of different terms related to attachement identifier and destination. The remaining keys appear to still
        /// be unique enough.
        /// 
        /// 2) Needed to merge N files instead of just two.
        /// 
        /// 3) The store retrieval was failing on _outlookNameSpace.Stores[PSTName] so I changed the search to be based on the FilePath property containing the to be merged PST file name. The original
        /// PST1 or PST2 PSTName lookups did not make sense to me. This could potential be a regression between outlook version. Not totally sure. These changes were made working with:
        /// Microsoft Office Professional Plus 2010.
        /// 
        /// Example CLI Invocation:
        /// C:\dev\EmailMerge\MergeProgram\bin\Release>MergeProgram.exe C:\dev\EmailMerge\working\merge.pst C:\dev\EmailMerge\working\archive_test\
        /// 
        /// </summary>
        /// <param name="args">Inputs in order:
        /// 1- merge filename
        /// 2- directory to search for psts to merge. Will not search recursive.

        static void Main(string[] args)
        {
            try
            {
                if (!args.Any() || args.Count() < 2) throw new ArgumentException("You must provide at least a merge filename and a search directory to merge psts for");
                Console.WriteLine("Loading and Merging PST Files");
        
                string mergeFile = args[0];
                string directoryToMerge = args[1];

                Console.WriteLine("File to Merge into:" + mergeFile);
                Console.WriteLine("Search Directory:" + directoryToMerge);

                PSTFile.MergePSTFiles(directoryToMerge, mergeFile, null,null,false,null);

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
