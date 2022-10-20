using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ExcelToVcard;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        InitializeComponent();
    }

    private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            OpenFileDialog openFileDialog = new() { Filter = "Excel files (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog(this) == false)
                return;

            this.Cursor = Cursors.Wait;

            using var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = Encoding.GetEncoding(1252) });

            var result = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = tableReader => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            foreach (var name in new[] { "STAND SUPPORT", "VISITORS" })
            {
                var table = result.Tables[name]?.AsEnumerable() ?? throw new System.Exception($"Table {name} not found");

                var vcards =
                    table
                    .Where(row => !String.IsNullOrEmpty(row["NAME"].ToString()))
                    .Select(row => $@"BEGIN:VCARD\nVERSION:3.0\nN;CHARSET=UTF-8:{row["SURNAME"].FixChars()};{row["NAME"].FixChars()}\nFN;CHARSET=UTF-8:{row["NAME"].FixChars()} {row["SURNAME"].FixChars()}\nORG:{row["COMPANY"].FixChars()}\nTITLE:{row["JOB TITLE"].FixChars()}\nTEL;CELL:{row["PHONE NUMBER"].FixChars().CheckPhone()}\nADR;WORK:;;{row["STREET"].FixChars()};{row["City"].FixChars()};{row["POSTCODE"].FixChars()};{row["COUNTRY"].FixChars()}\nURL:{row["WWW"].FixChars().FixUrl()}\nEMAIL;WORK;INTERNET:{row["E-MAIL"].FixChars()}\nnEND:VCARD")
                    .ToList();

                var path = Path.Combine(
                        Path.GetDirectoryName(openFileDialog.FileName) ?? throw new Exception("Directory not found"),
                        Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_" + name.ToLower() + ".txt");

                File.WriteAllText(path, "#QRCodes\n", System.Text.Encoding.Unicode);
                File.AppendAllLines(path, vcards, System.Text.Encoding.Unicode);
            }

            MessageBox.Show("Export completed", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            this.Cursor = Cursors.Arrow;
        }
    }

    private void BtnIdentifyFiles_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Choose folder to scan for vCard files",
            UseDescriptionForTitle = true,
            SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + Path.DirectorySeparatorChar,
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            return;

        foreach (var filename in Directory.GetFiles(dialog.SelectedPath, "*.jpg"))
        {

        }
    }
}
