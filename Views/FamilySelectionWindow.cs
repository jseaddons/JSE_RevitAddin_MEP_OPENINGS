using System.Windows;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    public partial class FamilySelectionWindow : Window
    {
        public FamilySelectionViewModel ViewModel { get; }

        public FamilySelectionWindow(FamilySelectionViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
            DataContext = ViewModel;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel.SelectedFamily != null)
                DialogResult = true;
            else
                MessageBox.Show("Please select a family.");
        }
    }
}
