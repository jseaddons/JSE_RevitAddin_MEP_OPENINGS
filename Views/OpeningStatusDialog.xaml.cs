using System.Windows;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Interaction logic for OpeningStatusDialog.xaml
    /// </summary>
    public partial class OpeningStatusDialog : Window
    {
        public OpeningStatusDialog()
        {
            InitializeComponent();
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}


