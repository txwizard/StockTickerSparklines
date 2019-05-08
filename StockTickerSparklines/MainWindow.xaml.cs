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

            txtSearchString.Focus ( );
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
                    ShowSearchResults (
                        tickerEngine ,
                        symbols );

                    this.txtMessage.Text = @"Symbols for you have I!";
                    cmdPruneSelections.IsEnabled = true;
                    cmdResetForm.IsEnabled = true;
                }   // TRUE (anticpated outcome) block, if ( symbols != null )
                else
                {
                    this.txtMessage.Text = tickerEngine.Message;
                }   // FALSE (unanticpated outcome) block, if ( symbols != null )
            }
            catch ( Exception ex )
            {
                ReportException ( ex );
            }   // catch ( Exception ex )
        }   // private void CmdSearch_Click event delegate


        private void CmdPruneSelections_Click ( object sender , RoutedEventArgs e )
        {
            string strKeepThisRow = Properties.Resources.XLS_ROW_DISP_KEEP;     // Save trips to the string table.
            bool fNewTopRow = false;
            List<KeptRow> keptRows = new List<KeptRow> ( __intAbsLastRow );

            for ( int intCurrAbsRow = __intAbsLastRow ;
                      intCurrAbsRow > ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intCurrAbsRow-- )
            {   // Work the grid from the bottom up.
                if ( xlWork.ActiveSheet.Cells [ intCurrAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].Value.Equals ( strKeepThisRow ) )
                {
                    keptRows.Add ( new KeptRow ( intCurrAbsRow ) );

                    if ( fNewTopRow )
                    {
                        for ( int intRowIndex = __intAbsLastRow - ArrayInfo.NEXT_INDEX ;
                                  intRowIndex > intCurrAbsRow ;
                                  intRowIndex-- )
                        {   // Work from the top down.
                            ClearPopulatedCellsInRow ( intRowIndex );
                        }   // for ( int intRowIndex = __intAbsLastRow - ArrayInfo.NEXT_INDEX ; intRowIndex > intCurrAbsRow ; intRowIndex-- )
                    }   // TRUE (Rows to keep lie above the current row.) block, if ( fNewTopRow )
                    else
                    {
                        for ( int intRowIndex = __intAbsLastRow ;
                                  intRowIndex > intCurrAbsRow ;
                                  intRowIndex-- )
                        {
                            ClearPopulatedCellsInRow ( intRowIndex );
                        }   // for ( int intRowIndex = __intAbsLastRow ; intRowIndex > intCurrAbsRow ; intRowIndex-- )

                        fNewTopRow = true;
                    }   // FALSE (Working form the bottom, this is the first row marked for retention.) block, if ( fNewTopRow )

                    __intAbsLastRow = intCurrAbsRow;
                }   // if ( xlWork.ActiveSheet.Cells [ intCurrAbsRow , ArrayInfo.ARRAY_FIRST_ELEMENT ].Value.Equals ( strKeepThisRow ) )
            }   // for ( int intCurrAbsRow = __intAbsLastRow ; intCurrAbsRow > ArrayInfo.ARRAY_FIRST_ELEMENT ; intCurrAbsRow-- )

            //  ----------------------------------------------------------------
            //  Process the last group of discarded rows.
            //  ----------------------------------------------------------------

            if ( fNewTopRow )
            {
                for ( int intRowIndex = __intAbsLastRow - ArrayInfo.NEXT_INDEX ;
                          intRowIndex > ArrayInfo.ARRAY_FIRST_ELEMENT ;
                          intRowIndex-- )
                {   // Work from the top down.
                    ClearPopulatedCellsInRow ( intRowIndex );
                }   // for ( int intRowIndex = __intAbsLastRow - ArrayInfo.NEXT_INDEX ; intRowIndex > ArrayInfo.ARRAY_FIRST_ELEMENT ; intRowIndex-- )
            }   // if ( fNewTopRow )

            if ( keptRows.Count > ListInfo.LIST_IS_EMPTY )
            {   // The list issn't empty.
                if ( keptRows [ ArrayInfo.ARRAY_FIRST_ELEMENT ].RowIndex >= keptRows.Count )
                {   // Since the index of the first item kept is greater than the length of the list, at least one item must be moved.
                    keptRows.Sort ( );

                    for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                              intRowIndex < keptRows.Count ;
                              intRowIndex++ )
                    {
                        MovePopulatedRow (
                            keptRows [ intRowIndex ].RowIndex ,
                            ArrayInfo.OrdinalFromIndex ( intRowIndex ) ,
                            __intAbsLastCol ,
                            xlWork.ActiveSheet.Cells );
                    }   // for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intRowIndex < keptRows.Count ; intRowIndex++ )
                }   // if ( keptRows [ ArrayInfo.ARRAY_FIRST_ELEMENT ].RowIndex >= keptRows.Count )

                __intAbsLastRow = keptRows.Count;
            }   // TRUE (anticipated outcome) block, if ( keptRows.Count > ListInfo.LIST_IS_EMPTY )
            else
            {
                __intAbsLastRow = ArrayInfo.ARRAY_SECOND_ELEMENT;
            }   // FALSE (unanticipated outcome) block, if ( keptRows.Count > ListInfo.LIST_IS_EMPTY )

            cmdGetHistory.IsEnabled = true;
            cmdPruneSelections.IsEnabled = false;
            txtMessage.Text = Properties.Resources.MSG_LIST_PRUNED;
        }   // CmdPruneSelections_Click event delegate


        private void CmdGetHistory_Click ( object sender , RoutedEventArgs e )
        {

        }   // CmdGetHistory_Click event delegate


        private void CmdResetForm_Click ( object sender , RoutedEventArgs e )
        {
            if ( MessageBox.Show ( Properties.Resources.MSG_ARE_YOU_SURE ,
                                   Title ,
                                   MessageBoxButton.YesNo ,
                                   MessageBoxImage.Question ) == MessageBoxResult.Yes )
            {
                txtSearchString.Text = SpecialStrings.EMPTY_STRING;

                for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                          intRowIndex <= __intAbsLastRow ;
                          intRowIndex++ )
                {
                    ClearPopulatedCellsInRow ( intRowIndex );
                }   // for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intRowIndex <= __intAbsLastRow ; intRowIndex++ )

                __intAbsLastRow = ArrayInfo.ARRAY_INVALID_INDEX;
                __intAbsLastCol = ArrayInfo.ARRAY_INVALID_INDEX;

                cmdGetHistory.IsEnabled = false;
                cmdPruneSelections.IsEnabled = false;
                cmdResetForm.IsEnabled = false;
                txtMessage.Text = Properties.Resources.MSG_ENTER_SEARCH_STRING;

                txtSearchString.Focus ( );
            }   // if ( MessageBox.Show ( Properties.Resources.MSG_ARE_YOU_SURE , Title , MessageBoxButton.YesNo , MessageBoxImage.Question ) == MessageBoxResult.Yes )
        }   // CmdResetForm_Click event delegate


        private void ClearPopulatedCellsInRow ( int pintRowIndex )
        {
            for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intColIndex <= __intAbsLastCol ;
                      intColIndex++ )
            {
                xlWork.ActiveSheet.Cells [ pintRowIndex , intColIndex ].Value = null;
            }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex <= __intAbsLastCol ; intColIndex++ )
        }   // ClearPopulatedCellsInRow



        private void MovePopulatedRow (
            int pintSourceRowIndex ,
            int pintDestinationRowIndex ,
            int pintAbsLastCol ,
            Cells pcells )
        {
            if ( pintSourceRowIndex > pintDestinationRowIndex )
            {
                for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                          intColIndex <= pintAbsLastCol ;
                          intColIndex++ )
                {
                    pcells [ pintDestinationRowIndex , intColIndex ].Value = pcells [ pintSourceRowIndex , intColIndex ].Value;
                    pcells [ pintSourceRowIndex , intColIndex ].Value = null;
                }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex <= pintAbsLastCol ; intColIndex++ )
            }   // if ( pintSourceRowIndex > pintDestinationRowIndex )
        }   // private void MovePopulatedRow


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

            //  ----------------------------------------------------------------
            //  Update the properites that maintain the address of the absolute 
            //  last row and column in the worksheet grid.
            //  ----------------------------------------------------------------

            if ( intAbsRow > __intAbsLastRow )
            {
                __intAbsLastRow = intAbsRow;
            }   // if ( intAbsRow > __intAbsLastRow )

            if ( intAbsCol > __intAbsLastCol )
            {
                __intAbsLastCol = intAbsCol;
            }   // if ( intAbsRow > __intAbsLastRow )
        }   // PopulateRowFromSearchResult


        private void ReportException ( Exception pexAllKinds )
        {
            MessageBox.Show (
                pexAllKinds.Message ,
                Title ,
                MessageBoxButton.OK ,
                MessageBoxImage.Exclamation );
            this.txtMessage.Text = pexAllKinds.Message;
            TraceLogger.WriteWithBothTimesLabeledLocalFirst (
                string.Format (
                    Properties.Resources.ERRMSG_WINMAIN_EXCEPTION ,
                    new string [ ]
                    {
                            pexAllKinds.GetType ( ).FullName ,                  // Format Item 0: An (0) Exception arose. 
                            pexAllKinds.Message ,                               // Format Item 1: Message             = {1}
                            pexAllKinds.TargetSite.Name ,                       // Format Item 2: TargetSite (Method) = {2}
                            pexAllKinds.Source ,                                // Format Item 3: Source (Assembly)   = {3}
                            pexAllKinds.StackTrace ,                            // Format Item 4: StackTrace          = {4}
                            Environment.NewLine                                 // Format Item 5: Platform-defined newline
                    } ) );
        }   // ReportException


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
                  intColIndex < __intAbsLastCol ;
                  intColIndex++ )
            {
                this.xlWork.View.AutoFitColumn ( intColIndex );
            }   // // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex < symbolInfos.Length ; intColIndex++ )

            this.xlWork.ActiveSheet.Protect = true;
        }   // private void ShowSearchResults


        private int __intAbsLastRow = ArrayInfo.ARRAY_INVALID_INDEX;
        private int __intAbsLastCol = ArrayInfo.ARRAY_INVALID_INDEX;


        private class KeptRow :IComparable<KeptRow>
        {
            private KeptRow ( )
            {
            }   // KeptRow constructor 1 of 2 is marked private, to force construction of initialized instances,


            public KeptRow ( int pintRowIndex )
            {
                RowIndex = pintRowIndex;
            }   // KeptRow constructor 2 of 2 is marked public, and its argument initializes the sole property of the class.


            public int RowIndex
            {
                get; private set;
            }   // public int RowIndex, a read-only property set by the only visible constructor


            public override bool Equals ( object obj )
            {
                KeptRow kept = obj as KeptRow;
                return RowIndex.Equals ( Invert ( kept.RowIndex ) );
            }   // public override bool Equals


            public override int GetHashCode ( )
            {
                return Invert ( RowIndex ).GetHashCode ( );
            }   // public override int GetHashCode


            public override string ToString ( )
            {
                return RowIndex.ToString ( );
            }   // public override string ToString


            int IComparable<KeptRow>.CompareTo ( KeptRow other )
            {
                return RowIndex.CompareTo ( Invert ( other.RowIndex ) );
            }   // int IComparable<KeptRow>.CompareTo


            private static int Invert ( int pintInput )
            {
                return pintInput * MagicNumbers.MINUS_ONE;
            }   // private static int Invert
        }   // private class KeptRow
    }   // public partial class MainWindow
}   // partial namespace StockTickerSparklines