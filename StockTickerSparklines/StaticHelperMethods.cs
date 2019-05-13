using System;

using Microsoft.Win32;

using GrapeCity.Windows.SpreadSheet.Data;

namespace StockTickerSparklines
{
    internal sealed class StaticHelperMethods
    {
        internal static string SaveWorkAsExcelWorkbook (
            GrapeCity.Windows.SpreadSheet.UI.GcSpreadSheet xlWork ,
            MainWindow phwndMain )
        {
            SaveFileDialog dialog = new SaveFileDialog ( );

            dialog.DefaultExt = @".xlsx";
            dialog.Filter = @"Microsoft Excel Documents|*" + dialog.DefaultExt;

            dialog.AddExtension = true;
            dialog.CheckPathExists = true;

            if ( ( bool ) dialog.ShowDialog ( phwndMain ) )
            {
                xlWork.SaveExcel (
                    dialog.FileName ,
                    ExcelFileFormat.XLSX ,
                    ExcelSaveFlags.NoFlagsSet );
                return dialog.FileName;
            }   // TRUE (anticipated outcome) block, if ( ( bool ) dialog.ShowDialog ( phwndMain ) )
            else
            {
                return null;
            }   // FALSE (unanticipated outcome) block, if ( ( bool ) dialog.ShowDialog ( phwndMain ) )
        }   // internal static void SaveWorkAsExcelWorkbook
    }   // internal sealed class StaticHelperMethods
}   // partial namespace StockTickerSparklines