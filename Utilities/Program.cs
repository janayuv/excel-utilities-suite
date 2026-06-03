using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace utilities
{
    static class Program
    {
        // The Main entry point for the application
        [STAThread]
        static void Main()
        {
            // Set visual styles and enable text rendering for the application
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Start the application with the main form
            Application.Run(new SequenceForm());  // Replace 'SequenceForm' with the name of your main form
        }
    }
}
