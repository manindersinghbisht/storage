using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Diagnostics.PerformanceData;

namespace MultiLocationZip
{
    public class MultipleLocationFileOrFolderToZip : NativeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Inset Multiple file or folder path in ',' seprated")]
        public InArgument<string> InputFilePaths { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Insert the path where you want to save the zip file")]
        public InArgument<string> DestinationPath { get; set; }


        [Category("Input")]
        [Description("if value of this field is empty then the default ZIP file name will be <todays date>__CompressFile.zip")]
        public InArgument<string> ZipFolderName { get; set; }
        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
        }


        protected override void Execute(NativeActivityContext context)
        {
            var todayDate = DateTime.UtcNow.Date;
            var in_FilePaths = InputFilePaths.Get(context);
            var in_DestPath = DestinationPath.Get(context);
//            var fileName = "";
            var in_ZipFolderName = ZipFolderName.Get(context);

            if (in_ZipFolderName.Equals("")) { in_ZipFolderName = todayDate.ToString("dd.MM.yyyy")+"_CompressFile"; }
            Directory.CreateDirectory(in_DestPath + "\\"+"TempFile");

            var patharray = in_FilePaths.Split(',');

            CopyFolderFile(patharray, in_DestPath+"\\TempFile");

            /*
            foreach(string file in patharray)
            {
 
                fileName = Path.GetFileName(file);

                FileAttributes fa = File.GetAttributes(file);

                if ((fa & FileAttributes.Directory) == FileAttributes.Directory) {
                    

                }

                File.Copy(file, in_DestPath  + "\\" + "TempFile\\" +fileName,true);

        }
        */


        ZipFile.CreateFromDirectory(in_DestPath + "\\" + "TempFile", in_DestPath + "\\"+in_ZipFolderName+".zip");
            Directory.Delete(in_DestPath + "\\" + "TempFile",true);
        }


        public static void CopyFolderFile(string[] folderdirectory, string in_DestPath ) {


            foreach (string item in folderdirectory) {

                FileAttributes fa = File.GetAttributes(item);

                if ((fa & FileAttributes.Directory) == FileAttributes.Directory)
                {

                    var directoryName = Path.GetFileName(item);

                    Directory.CreateDirectory(in_DestPath + "\\" + directoryName);

                    var childFiles = Directory.GetFiles(item);

                    CopyFolderFile(childFiles,in_DestPath+"\\"+directoryName);
                }

                else {

                    var fileName = Path.GetFileName(item);

                    File.Copy(item, in_DestPath + "\\"+ fileName, true);

                };

            }
        
        }
    }
}
