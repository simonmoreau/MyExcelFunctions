using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelFunctions
{

    public class BimsyncFolder
    {
        public List<BimsyncFolder> Children = new List<BimsyncFolder>();
        public BimsyncFolder Parent { get; set; }
        public MyFolder AssociatedFolder{ get; set; }
    }

    public class Folder
    {
        public Folder(BimsyncFolder bimsyncFolder)
        {
            this.name = bimsyncFolder.AssociatedFolder.Name;

            foreach (BimsyncFolder children in bimsyncFolder.Children)
            {
                if (this.folders == null) {this.folders = new List<Folder>(); }
                this.folders.Add(new Folder(children));
            }
        }

        public string name { get; set; }
        public List<Folder> folders { get; set; }
    }


    public class MyFolder
    {
        public MyFolder(string parentId, string name)
        {
            this.ParentID = parentId;
            this.Name = name;
        }
        public string ParentID { get; set; }
        public string ID { get { return this.ParentID + @"\" + this.Name; } }
        public string Name { get; set; }
    }
}
