using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using GrapeCity.Windows.SpreadSheet.Data;

using WizardWrx;
using WizardWrx.Core;

namespace StockTickerSparklines
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow ( )
        {
            InitializeComponent ( );
            this.Title = System.Reflection.Assembly.GetExecutingAssembly ( ).GetName ( ).Name;
        }   // public MainWindow default constructor


        private void Window_ContentRendered ( object sender , EventArgs e )
        {
            //  ----------------------------------------------------------------
            //  Until the ContentRendered event arises, you cannot set the focus
            //  on a control that is part of said content. While simple windows
            //  can get away with setting the focus on a control in their
            //  constructors, the ContentRendered event delegate is a better
            //  choice, since it is the last event that is raised before control
            //  is relinquished to the user.
            //  ----------------------------------------------------------------

            this.txtSearchString.Focus ( );
        }   // private void Window_ContentRendered event delegate


        private void TxtSearchString_TextChanged ( object sender , TextChangedEventArgs e )
        {
            TextBox textBox = sender as TextBox;

            this.cmdSearch.IsEnabled = ( textBox.Text.Length > ListInfo.EMPTY_STRING_LENGTH );
        }   // private void TxtSearchString_TextChanged event delegate
 
 
        private void CmdSearch_Click ( object sender , RoutedEventArgs e )
        {
            try
            {
                StockTickerEngine tickerEngine = StockTickerEngine.GetTheSingleInstance ( );
                TickerSymbolMatches symbols = tickerEngine.Search ( txtSearchString.Text );

                if ( symbols != null )
                {
                    ShowSearchResults ( tickerEngine , symbols );

                    this.txtMessage.Text = @"Symbols for you have I!";
                }   // TRUE (anticpated outcome) block, if ( symbols != null )
                else
                {
                    this.txtMessage.Text = tickerEngine.Message;
                }   // FALSE (unanticpated outcome) block, if ( symbols != null )
            }
            catch ( Exception ex )
            {
                MessageBox.Show (
                    ex.Message ,
                    Title ,
                    MessageBoxButton.OK ,
                    MessageBoxImage.Exclamation );
                this.txtMessage.Text = ex.Message;
                TraceLogger.WriteWithBothTimesLabeledLocalFirst (
                    string.Format (
                        Properties.Resources.ERRMSG_WINMAIN_EXCEPTION ,
                        new string [ ]
                        {
                            ex.GetType ( ).FullName ,       // Format Item 0: An (0) Exception arose. 
                            ex.Message ,                    // Format Item 1: Message             = {1}
                            ex.TargetSite.Name ,            // Format Item 2: TargetSite (Method) = {2}
                            ex.Source ,                     // Format Item 3: Source (Assembly)   = {3}
                            ex.StackTrace ,                 // Format Item 4: StackTrace          = {4}
                            Environment.NewLine             // Format Item 5: Platform-defined newline
                        } ) );
            }   // catch ( Exception ex )
        }   // private void CmdSearch_Click event delegate


        private void ShowSearchResults (
            StockTickerEngine ptickerEngine ,
            TickerSymbolMatches psymbols )
        {
            //  ----------------------------------------------------------------
            //  Put IntelliSense to work for me; defining these as constants
            //  causes both IntelliSense and the C# compiler to flag things that
            //  would otherwise by typographical errors that would go undetected
            //  until runtime.
            //  ----------------------------------------------------------------

            const string FONT_FAMILY_TAHOMA = @"Tahoma";
            const string THEME_NAME_MY_THEME = @"thmMyTheme";
            const string THEME_TO_APPLY_HEADING = @"Headings";
            const string THEME_TO_APPLY_BODY = @"Body";

            SymbolInfo [ ] symbolInfos = ptickerEngine.GetSymbolInfos ( );

            SpreadTheme theme = new SpreadTheme ( THEME_NAME_MY_THEME );

            theme.BodyFontName = FONT_FAMILY_TAHOMA;
            theme.HeadingFontName = FONT_FAMILY_TAHOMA;
            theme.Colors.Accent1 = Color.FromArgb ( 255 , 255 , 0 , 0 );
            theme.Colors.Accent2 = Color.FromArgb ( 255 , 0 , 255 , 0 );

            var validChoices = DataValidator.CreateListValidator (
                string.Concat (
                    Properties.Resources.XLS_ROW_DISP_DISCARD ,
                    SpecialCharacters.COMMA ,
                    Properties.Resources.XLS_ROW_DISP_KEEP ) );

            //  --------------------------------------------------------
            //  The first column is reserved for marking selections to
            //  be kept. Whereas worksheet cells in the Microsoft Excel
            //  object model have ordinal numbers that start at one, the
            //  GrapeCity control wisely treats them as a first class
            //  array.
            //  --------------------------------------------------------

            this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , ArrayInfo.ARRAY_FIRST_ELEMENT ].Value = Properties.Resources.XLS_R1_C1_LABEL;
            this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , ArrayInfo.ARRAY_FIRST_ELEMENT ].FontTheme = THEME_TO_APPLY_HEADING;
            this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , ArrayInfo.ARRAY_FIRST_ELEMENT ].FontWeight = FontWeights.Bold;
            this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , ArrayInfo.ARRAY_FIRST_ELEMENT ].Locked = false;

            //  --------------------------------------------------------
            //  Label the next 9 columns, set the font weight to bold,
            //  and protect the cells.
            //  --------------------------------------------------------

            for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intColIndex < symbolInfos.Length ;
                      intColIndex++ )
            {
                int intAbsCol = ArrayInfo.OrdinalFromIndex ( intColIndex );

                this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , intAbsCol ].Value = symbolInfos [ intColIndex ].Label;
                this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , intAbsCol ].FontWeight = FontWeights.Bold;
                this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , intAbsCol ].Locked = true;
            }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex < symbolInfos.Length ; intColIndex++ )

            //  --------------------------------------------------------
            //  Populate the detail rows.
            //  --------------------------------------------------------

            for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intRowIndex < psymbols.bestMatches.Length ;
                      intRowIndex++ )
            {
                int intAbsRow = ArrayInfo.OrdinalFromIndex ( intRowIndex );

                this.xlWork.ActiveSheet.Cells [ intAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].Value = Properties.Resources.XLS_ROW_DISP_DISCARD;
                this.xlWork.ActiveSheet.Cells [ intAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].FontTheme = THEME_TO_APPLY_BODY;
                this.xlWork.ActiveSheet.Cells [ intAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].FontWeight = FontWeights.Normal;
                this.xlWork.ActiveSheet.Cells [ intAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].Locked = false;
                this.xlWork.ActiveSheet.Cells [ intAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].DataValidator = validChoices;

                for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intColIndex < symbolInfos.Length ;
                      intColIndex++ )
                {
                    PopulateRowFromSearchResult (
                        psymbols ,
                        symbolInfos ,
                        intRowIndex ,
                        intColIndex );
                }   // // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex < symbolInfos.Length ; intColIndex++ )
            }   // for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intRowIndex < symbols.bestMatches.Length ; intRowIndex++ )

            //  --------------------------------------------------------
            //  Auto-fit the column widths.
            //  --------------------------------------------------------

            for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                  intColIndex < symbolInfos.Length ;
                  intColIndex++ )
            {
                this.xlWork.View.AutoFitColumn ( intColIndex );
            }   // // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex < symbolInfos.Length ; intColIndex++ )

            this.xlWork.ActiveSheet.Protect = true;
        }   // private void ShowSearchResults


        private void PopulateRowFromSearchResult (
            TickerSymbolMatches psymbols ,
            SymbolInfo [ ] paSymbolInfos ,
            int pintRowIndex ,
            int pintColIndex )
        {
            //  ----------------------------------------------------------------
            //  Since the first column is reserved for the disposition choices,
            //  subsequent columns are shifted right by one. Likewise, since the
            //  first row is reserved for labels, the row index is also shifted,
            //  down by one.
            //  ----------------------------------------------------------------

            const int COL_IDX_SYMBOL = 0;
            const int COL_IDX_NAME = 1;
            const int COL_IDX_TYPE = 2;
            const int COL_IDX_REGION = 3;
            const int COL_IDX_MARKETOPEN = 4;
            const int COL_IDX_MARKETCLOSE = 5;
            const int COL_IDX_TIMEZONE = 6;
            const int COL_IDX_CURRENCY = 7;
            const int COL_IDX_MATCHSCORE = 8;

            string strCellValue = null;

            //  ------------------------------------------------
            //  The column index determines which of the nine
            //  values in the BestMatches array goes into the
            //  current column. As a matter of habit, I code for
            //  the default case, even when it should never
            //  happen.
            //  ------------------------------------------------

            switch ( pintColIndex )
            {
                case COL_IDX_SYMBOL:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._1symbol;
                    break;
                case COL_IDX_NAME:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._2name;
                    break;
                case COL_IDX_TYPE:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._3type;
                    break;
                case COL_IDX_REGION:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._4region;
                    break;
                case COL_IDX_MARKETOPEN:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._5marketOpen;
                    break;
                case COL_IDX_MARKETCLOSE:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._6marketClose;
                    break;
                case COL_IDX_TIMEZONE:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._7timezone;
                    break;
                case COL_IDX_CURRENCY:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._8currency;
                    break;
                case COL_IDX_MATCHSCORE:
                    strCellValue = psymbols.bestMatches [ pintRowIndex ]._9matchScore;
                    break;
                default:
                    throw new InvalidOperationException ( string.Format (
                        Properties.Resources.TPL_INTERNAL_ERROR_001 ,           // Format control string, pulled the language-neutral resources embedded into the assembly
                        pintColIndex ,                                          // Format Item 0: intColIndex has an invalid value of {0}.
                        paSymbolInfos.Length ) );                               // Format Item 1: Its value must always be less than {1}.
            }   // switch ( intColIndex )

            //  ----------------------------------------------------------------
            //  The correct value is in strCellValue; use it to set the Value
            //  property, then protect the cell. Since they are used twice, both
            //  offsets are computed and stored in local scratch variables.
            //  ----------------------------------------------------------------

            int intAbsRow = ArrayInfo.OrdinalFromIndex ( pintRowIndex );
            int intAbsCol = ArrayInfo.OrdinalFromIndex ( pintColIndex );

            this.xlWork.ActiveSheet.Cells [ intAbsRow , intAbsCol ].Value = strCellValue;
            this.xlWork.ActiveSheet.Cells [ intAbsRow , intAbsCol ].Locked = true;
        }   // PopulateRowFromSearchResult
    }   // public partial class MainWindow
}   // partial namespace StockTickerSparklines