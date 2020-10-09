using System;
using System.Diagnostics;
using TTExcelGenerator;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace ExcelTest
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void Grid_DragOver(object sender, DragEventArgs e)
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
        }

        private async void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.DataView.Contains(StandardDataFormats.StorageItems))
            {
                var items = await e.DataView.GetStorageItemsAsync();
                if (items.Count > 0)
                {
                    var storageFile = items[0] as StorageFile;
                    if (storageFile != null)
                    {
                        Workbook wk = new Workbook();
                        wk.Open(storageFile);

                        Worksheet wksheet1 = wk.GetSheetAt(0);
                        for (int i = 0; i < wksheet1.Rows.Count; i++)
                        {
                            Debug.WriteLine(wksheet1.Rows[i].Cells.Count);
                            for (int j = 0; j < wksheet1.Rows[i].Cells.Count; j++)
                            {
                                Debug.WriteLine(wksheet1.Rows[i].Cells[j].Value);
                            }
                        }
                    }
                }
            }
        }
    }
}
