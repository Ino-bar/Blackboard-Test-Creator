using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Blackboard_Test_Creator
{
    class TestCreator
    {
        public string savePath = Form1.TestFilePath;
        public void CreateFile(string FilePath, string FileName)
        {
            File.Create(FilePath + "\\" + FileName);
        }
        public void CreateBBPackage()
        {
            CreateFile(savePath, ".bb-package-info");
        }
        public void Createimsmanifest()
        {
            CreateFile(savePath, "imsmanifest.xml");
        }
        public void Createres00001()
        {
            CreateFile(savePath, "res00001.dat");
        }
        public void Createres00002()
        {
            CreateFile(savePath, "res00002.dat");
        }
        public void Createres00003to5()
        {
            CreateFile(savePath, "res00003.dat");
            CreateFile(savePath, "res00004.dat");
            CreateFile(savePath, "res00005.dat");
        }
    }
}
