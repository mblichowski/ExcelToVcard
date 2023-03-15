using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using vCardLib.Deserializers;
using BarcodeReader = ZXing.Presentation.BarcodeReader;

namespace ExcelToVcard;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private const string _subDirectory = "identified vCards";
    private const int _qrCodeImagePixelHeightDefault = 6000;
    private readonly int? _qrCodeImagePixelHeight;

    private readonly BarcodeReader _reader = new BarcodeReader();

    public MainWindow()
    {
        _qrCodeImagePixelHeight =
            int.TryParse(ConfigurationManager.AppSettings["QrCodeImagePixelHeight"], out int pixelHeight) ?
            pixelHeight :
            _qrCodeImagePixelHeightDefault;

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        InitializeComponent();
    }

    private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            OpenFileDialog openFileDialog = new() { Filter = "Excel files (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog(this) == false)
                return;

            this.Cursor = Cursors.Wait;
            this.progressBar.Visibility = Visibility.Visible;

            await Task.Run(() => CreateVCardTemplateFilesAsync(openFileDialog));

            MessageBox.Show("Export completed", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            this.Cursor = Cursors.Arrow;
            this.progressBar.Visibility = Visibility.Hidden;
        }

        static void CreateVCardTemplateFilesAsync(OpenFileDialog openFileDialog)
        {
            FileStream stream;
            IExcelDataReader reader;

            stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
            reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = Encoding.GetEncoding(1252) });
            var result = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = tableReader => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            var sheets = result.Tables.OfType<DataTable>();
            if (!sheets.Any())
                throw new Exception($"No sheets found in the given Excel workbook.");

            foreach (var sheet in sheets)
            {
                var vcards =
                    sheet
                    .AsEnumerable()
                    .Where(row => !String.IsNullOrEmpty(row["NAME"].ToString()))
                    .Select(row => $@"BEGIN:VCARD\nVERSION:3.0\nN;CHARSET=UTF-8:{row["SURNAME"].FixChars()};{row["NAME"].FixChars()}\nFN;CHARSET=UTF-8:{row["NAME"].FixChars()} {row["SURNAME"].FixChars()}\nORG:{row["COMPANY"].FixChars()}\nTITLE:{row["JOB TITLE"].FixChars()}\nTEL;CELL:{row["PHONE NUMBER"].FixChars().CheckPhone()}\nADR;WORK:;;{row["STREET"].FixChars()};{row["City"].FixChars()};{row["POSTCODE"].FixChars()};{row["COUNTRY"].FixChars()}\nURL:{row["WWW"].FixChars().FixUrl()}\nEMAIL;WORK;INTERNET:{row["E-MAIL"].FixChars()}\nEND:VCARD")
                    .ToList();

                var path = Path.Combine(
                        Path.GetDirectoryName(openFileDialog.FileName) ?? throw new Exception("Directory not found"),
                        Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_" + sheet.TableName.ToLower() + ".txt");

                File.WriteAllText(path, "#QRCodes\n", Encoding.Unicode);
                File.AppendAllLines(path, vcards, Encoding.Unicode);
            }
        }
    }

    private async void BtnIdentifyFiles_Click(object sender, RoutedEventArgs e)
    {
        string GetDestinationDirectory(string rootPath) => Path.Combine(rootPath, _subDirectory);

        try
        {
            using var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Choose folder to scan for vCard files",
                UseDescriptionForTitle = true,
                SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + Path.DirectorySeparatorChar,
            };

            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            this.Cursor = Cursors.Wait;
            this.progressBar.Visibility = Visibility.Visible;

            var b = chkAddSuffix.IsChecked ?? false;
            var failedFiles = await Task.Run(() => IdentifyvCards(dialog, b));
            if (failedFiles.Any())
                File.WriteAllLines(Path.Combine(GetDestinationDirectory(dialog.SelectedPath), "failed files.txt"), failedFiles);

            MessageBox.Show($"vCards identification completed.\n\nNew files written in: {GetDestinationDirectory(dialog.SelectedPath)}{(failedFiles.Any() ? "\n\nNot recognized files list written in the same directory." : "")}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            this.Cursor = Cursors.Arrow;
            this.progressBar.Visibility = Visibility.Hidden;
        }

        IEnumerable<string> IdentifyvCards(System.Windows.Forms.FolderBrowserDialog dialog, bool useDifferentDestNameTemplate)
        {
            List<string> failedFiles = new();

            Directory.CreateDirectory(GetDestinationDirectory(dialog.SelectedPath));
            foreach (var filename in Directory.GetFiles(dialog.SelectedPath, "*.jpg").OrderBy(_ => _))
            {
                BitmapImage bi = new();
                bi.BeginInit();
                bi.UriSource = new Uri(filename);
                bi.DecodePixelHeight = _qrCodeImagePixelHeight ?? _qrCodeImagePixelHeightDefault;
                bi.EndInit();

                var qrCodeContent = _reader.Decode(bi);

                if (qrCodeContent is null)
                {
                    failedFiles.Add(filename);
                    continue;
                }

                var vCard = Deserializer.FromString(qrCodeContent.Text).Single();

                var src = Path.Combine(dialog.SelectedPath, filename);
                var dest = Path.ChangeExtension(
                    path: Path.Combine(
                        GetDestinationDirectory(dialog.SelectedPath),
                        useDifferentDestNameTemplate ?
                        String.Join('_', vCard.CustomFields.Single(_ => !_.Key.Contains("FN")).Value.Trim().Split(";").Reverse()) + "_vCard" :
                        vCard.CustomFields.Single(_ => _.Key.Contains("FN")).Value),
                    extension: Path.GetExtension(filename));

                File.Copy(src, dest, overwrite: true);
            }

            return failedFiles;
        }
    }

}
