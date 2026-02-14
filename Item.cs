using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VideoEditor
{
    class Item
    {
        public FileType fileType
        {
            get; set;
        }
        public string path { get; set; }
        public string type { get; set; }
        public string fileName { get; set; }
        public int grid { get; set; }

        public double startPoint { get; set; }
        public double endPoint { get; set; }

        public double startVideo { get; set; }
        public double endVideo { get; set; }
        public double duration { get; set; }
        public Button button { get; set; }
        public int linkedItemId { get; set; } = -1;
        public float Volume { get; set; } = 1.0f;

    }
}
