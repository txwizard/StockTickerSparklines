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
        #region Private Object-scoped Enumerations and Symbolic Constants
        enum CellMovementDirection
        {
            Down ,
            Up
        }   // enum CellMovementDirection


        enum ProcessingState
        {
            Idle ,
            Searching,
            Pruning,
            GettingHistory,
            ResettingForm
        }   // enum ProcessingState


        const int COL_IDX_SYMBOL = 0;
        #endregion  // Private Object-scoped Enumerations and Symbolic Constants


        #region Constructor
        public MainWindow ( )
        {
            InitializeComponent ( );

            this.Title = System.Reflection.Assembly.GetExecutingAssembly ( ).GetName ( ).Name;
        }   // public MainWindow default constructor
        #endregion  // Constructor


        #region Event Delegates
        private void ActiveSheet_CellChanged ( object sender , CellChangedEventArgs e )
        {
            switch ( __enmProcessingState )
            {
                case ProcessingState.GettingHistory:
                case ProcessingState.Pruning:
                case ProcessingState.ResettingForm:
                case ProcessingState.Searching:
                    break;
                case ProcessingState.Idle:
                default:
                    if ( e.Column == COL_IDX_SYMBOL && e.PropertyName == Properties.Resources.XLS_PROPERTY_NAME_IS_VALUE )
                    {
                        try
                        {
                            Worksheet sheet = sender as Worksheet;              // Cast sender to a Worksheet so that its Cells array can be evaluated.

                            if ( ( string ) sheet.Cells [ e.Row , e.Column ].Value == Properties.Resources.XLS_ROW_DISP_KEEP )
                            {
                                cmdPruneSelections.IsEnabled = true;
                            }   // if ( ( string ) sheet.ActiveSheet.Cells [ e.Row , e.Column ].Value == Properties.Resources.XLS_ROW_DISP_KEEP )
                        }
                        catch ( Exception ex )
                        {
                            ReportException ( ex );
                        }
                    }   // if ( e.Column == COL_IDX_SYMBOL && e.PropertyName == Properties.Resources.XLS_PROPERTY_NAME_IS_VALUE )

                    break;  // case ProcessingState.Idle: and case default:
            }   // switch ( __enmProcessingState )
        }   // private void ActiveSheet_CellChanged


        private void CmdExportToExcel_Click ( object sender , RoutedEventArgs e )
        {
            Button button = sender as Button;

            string strExcelFileName = StaticHelperMethods.SaveWorkAsExcelWorkbook (
                xlWork ,
                this );

            if ( !string.IsNullOrEmpty ( strExcelFileName ) )
            {
                string strMessage = string.Format (
                    Properties.Resources.MSG_EXCEL_SAVED_AS ,
                    strExcelFileName );
                txtMessage.Text = strMessage;
                MessageBox.Show (
                    strMessage ,
                    Title ,
                    MessageBoxButton.OK ,
                    MessageBoxImage.Exclamation );
            }   // TRUE (anticipated outcome) block, if ( !string.IsNullOrEmpty ( strExcelFileName ) )
            else
            {
                txtMessage.Text = Properties.Resources.MSG_EXCEL_SAVE_ABORTED;
            }   // FALSE (unanticipated outcome) block, if ( !string.IsNullOrEmpty ( strExcelFileName ) )
        }   // private void CmdExportToExcel_Click


        private void CmdGetHistory_Click ( object sender , RoutedEventArgs e )
        {
            __enmProcessingState = ProcessingState.GettingHistory;

            GetHistoryForSelectedSymbols ( );

            cmdExportToExcel.IsEnabled = true;
            Button button = sender as Button;
            button.IsEnabled = false;

            __enmProcessingState = ProcessingState.Idle;
        }   // private void CmdGetHistory_Click event delegate


        private void CmdPruneSelections_Click ( object sender , RoutedEventArgs e )
        {
            __enmProcessingState = ProcessingState.Pruning;
            PruneSymbolsList ( );
            __enmProcessingState = ProcessingState.Idle;
        }   // CmdPruneSelections_Click event delegate


        private void CmdResetForm_Click ( object sender , RoutedEventArgs e )
        {
            if ( MessageBox.Show ( Properties.Resources.MSG_ARE_YOU_SURE ,
                                   Title ,
                                   MessageBoxButton.YesNo ,
                                   MessageBoxImage.Question ) == MessageBoxResult.Yes )
            {
                ResetForm ( );
            }   // if ( MessageBox.Show ( Properties.Resources.MSG_ARE_YOU_SURE , Title , MessageBoxButton.YesNo , MessageBoxImage.Question ) == MessageBoxResult.Yes )
        }   // private void CmdResetForm_Click event delegate


        private void CmdSearch_Click ( object sender , RoutedEventArgs e )
        {
            try
            {
                __enmProcessingState = ProcessingState.Searching;
                txtMessage.Text = Properties.Resources.MSG_SEARCH_UNDERWAY;

                StockTickerEngine tickerEngine = StockTickerEngine.GetTheSingleInstance ( );
                TickerSymbolMatches symbols = tickerEngine.Search ( txtSearchString.Text );

                if ( symbols != null )
                {
                    ShowSearchResults (
                        tickerEngine ,
                        symbols );

                    this.txtMessage.Text = Properties.Resources.MSG_SYMBOLS_FOUND;
                    cmdResetForm.IsEnabled = true;
                }   // TRUE (anticpated outcome) block, if ( symbols != null )
                else
                {
                    this.txtMessage.Text = tickerEngine.Message;
                }   // FALSE (unanticpated outcome) block, if ( symbols != null )

                __enmProcessingState = ProcessingState.Idle;
            }
            catch ( Exception ex )
            {
                ReportException ( ex );
            }   // catch ( Exception ex )
        }   // private void CmdSearch_Click event delegate


        private void TxtSearchString_TextChanged ( object sender , TextChangedEventArgs e )
        {
            TextBox textBox = sender as TextBox;

            this.cmdSearch.IsEnabled = ( textBox.Text.Length > ListInfo.EMPTY_STRING_LENGTH );
        }   // private void TxtSearchString_TextChanged event delegate


        private void Window_ContentRendered ( object sender , EventArgs e )
        {
            //  ----------------------------------------------------------------
            //  You cannot set the focus on a control until the content of which
            //  it is a part becomes the subject of a ContentRendered event. 
            //  While simple windows can get away with setting the focus on a
            //  control in their constructors, their ContentRendered event
            //  delegate is a more robust choice, since it is the last event 
            //  that is raised before control is relinquished to the user.
            //  ----------------------------------------------------------------

            txtSearchString.Focus ( );
            xlWork.ActiveSheet.CellChanged += this.ActiveSheet_CellChanged;
        }   // private void Window_ContentRendered event delegate
        #endregion  // Event Delegates


        #region Private Instance Worker Methods
        private string ApplyFixups_Pass_2 ( string pstrFixedUp_Pass_1 )
        {   // This method references and updates instance member __fIsFirstPass.
            const string TSD_LABEL_ANTE = "\"TimeSeriesDaily\": {";             // Ante: "TimeSeriesDaily": {
            const string TSD_LABEL_POST = "\"Time_Series_Daily\" : [";          // Post: "Time_Series_Daily": [

            const string END_BLOCK_ANTE = "}\n    }\n}";
            const string END_BLOCK_POST = "}\n    ]\n}";

            const int DOBULE_COUNTING_ADJUSTMENT = MagicNumbers.PLUS_ONE;       // Deduct one from the length to account for the first character occupying the position where copying begins.

            __fIsFirstPass = true;                                              // Re-initialize the First Pass flag.

            StringBuilder builder1 = new StringBuilder ( pstrFixedUp_Pass_1.Length * MagicNumbers.PLUS_TWO );

            builder1.Append (
                pstrFixedUp_Pass_1.Replace (
                    TSD_LABEL_ANTE ,
                    TSD_LABEL_POST ) );
            int intLastMatch = builder1.ToString ( ).IndexOf ( TSD_LABEL_POST )
                               + TSD_LABEL_POST.Length
                               - DOBULE_COUNTING_ADJUSTMENT;

            while ( intLastMatch > ListInfo.INDEXOF_NOT_FOUND )
            {
                intLastMatch = FixNextItem (
                    builder1 ,
                    intLastMatch );
            }   // while ( intLastMatch > ListInfo.INDEXOF_NOT_FOUND )

            //  ----------------------------------------------------------------
            //  Close the array by replacing the last French brace with a square
            //  bracket.
            //  ----------------------------------------------------------------

            builder1.Replace (
                END_BLOCK_ANTE ,
                END_BLOCK_POST );

            return builder1.ToString ( );
        }   // private string ApplyFixups_Pass_2


        private void ClearPopulatedCellsInRow ( int pintRowIndex )
        {
            for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intColIndex <= __intAbsLastCol ;
                      intColIndex++ )
            {
                xlWork.ActiveSheet.Cells [ pintRowIndex , intColIndex ].Value = null;
            }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex <= __intAbsLastCol ; intColIndex++ )
        }   // ClearPopulatedCellsInRow


        private void GetHistoryForSelectedSymbols ( )
        {
            const int STOCK_SYMBOL_COLUMN = 1;
            const int STOCK_SYMBOL_VIEWPORT_INDEX = 0;

            const int SPARKLINE_COLUMN_WIDTH = 100;
            const int SPARKLINE_VIEWPORT_WIDTH = SPARKLINE_COLUMN_WIDTH + 10;

            try
            {
                string [ ] astrTimeSeriesLabels = WizardWrx.EmbeddedTextFile.Readers.LoadTextFileFromEntryAssembly (
                    @"Time_Series_Daily.txt" );

                int [ ] aintIssueRows = MakeRoomForHistory (
                    xlWork.ActiveSheet ,
                    __intAbsLastRow ,
                    __intAbsLastCol ,
                    ArrayInfo.IndexFromOrdinal ( astrTimeSeriesLabels.Length ) );

                for ( int intCurrentIssue = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                          intCurrentIssue < aintIssueRows.Length ;
                          intCurrentIssue++ )
                {
                    int intCurrRow = aintIssueRows [ intCurrentIssue ];         // This value is used twice.
                    DailyTimeSeriesResponse tsrHistory = GetHistoryForSymbol (
                        xlWork.ActiveSheet.Cells [ intCurrRow , STOCK_SYMBOL_COLUMN ].Text ,
                        intCurrRow ,
                        __intAbsLastCol );

                    if ( tsrHistory != null )
                    {
                        //  ----------------------------------------------------
                        //  The value of intFirstDetailColumn is equal to the
                        //  sum of the current value of __intAbsLastCol, which
                        //  is the last column used to report the stock issue
                        //  query response, plus two, as follows:
                        //
                        //  1) The label column
                        //  2) The first daily time series column
                        //  ----------------------------------------------------

                        int intFirstDetailColumn = __intAbsLastCol + MagicNumbers.PLUS_TWO;
                        SparklineRowAssignments assignments = StoreTimeSeriesInWorksheet (
                            intCurrRow ,                                        // int               pintOriginRow        The row index of the first cell in the range.
                            __intAbsLastCol + ArrayInfo.NEXT_INDEX ,            // int               pintOriginCol        The column index of the first cell in the range.
                            astrTimeSeriesLabels ,                              // int [ ]           pastrRowLabels       The string array of labels for the rows
                            tsrHistory.Time_Series_Daily ,                      // Time_Series_Daily patsdTimeSeriesDaily The number of columns in the range.
                            xlWork.ActiveSheet );                               // Worksheet         pwsActiveSheet       The worksheet to fill

                        __intAbsLastRow = UpdateLastRow ( astrTimeSeriesLabels , intCurrRow , __intAbsLastRow );

                        xlWork.ActiveSheet.Columns [ assignments.SparklineColumnIndex ].Width = SPARKLINE_COLUMN_WIDTH;
                        xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , assignments.SparklineColumnIndex ].Value = Properties.Resources.COL_LBL_SPARKLINE_GRAPHS;

                        __intAbsLastCol = assignments.SparklineColumnIndex;

                        int [ ] aintSparklineRowAssignments = assignments.GetArrayOfSparklineRows ( );
                        CreateSparklines (
                            intFirstDetailColumn ,                              // int                                              pintFirstDetailColumn
                            tsrHistory.Time_Series_Daily.Length ,               // int                                              pintDetailColumnCount
                            assignments.SparklineColumnIndex ,                  // int                                              pintSparklineColumnIndex
                            aintSparklineRowAssignments ,                       // int [ ]                                          paintSparklineRowAssignments
                            xlWork );                                           // GrapeCity.Windows.SpreadSheet.UI.GcSpreadSheet   pxlWork
                    }   // if ( tsrHistory != null )
                }   // for ( int intCurrentIssue = ArrayInfo.ARRAY_FIRST_ELEMENT ; intCurrentIssue < aintIssueRows.Length ; intCurrentIssue++ )

                xlWork.AddColumnViewport (
                    STOCK_SYMBOL_VIEWPORT_INDEX ,
                    this.Width - SPARKLINE_VIEWPORT_WIDTH );

                //xlWork.AddColumnViewport (
                //    SPARKLINE_VIEWPORT_INDEX ,
                //    SPARKLINE_VIEWPORT_WIDTH );
                //xlWork.ColumnSplitBoxPolicy = GrapeCity.Windows.SpreadSheet.UI.SplitBoxPolicy.Always;
                //xlWork.SetViewportLeftColumn (
                //    SPARKLINE_VIEWPORT_INDEX ,
                //    SPARKLINE_VIEWPORT_LEFT_COLUMN_INDEX );
                //xlWork.Invalidate ( );

                //  ----------------------------------------------------------------
                //  Freeze the top row and the first and last columns.
                //  ----------------------------------------------------------------

                xlWork.ActiveSheet.FrozenRowCount = 1;
                xlWork.ActiveSheet.FrozenColumnCount = 1;
                xlWork.ActiveSheet.FrozenTrailingColumnCount = 1;
                xlWork.SetViewportLeftColumn (
                    ArrayInfo.IndexFromOrdinal (
                        xlWork.GetColumnViewportCount ( ) ) ,
                    __intAbsLastCol );
                xlWork.Invalidate ( );
                xlWork.UpdateLayout ( );
            }
            catch ( Exception ex )
            {
                ReportException ( ex );
            }   // catch ( Exception ex )
        }   // private void GetHistoryForSelectedSymbols


        private DailyTimeSeriesResponse GetHistoryForSymbol (
            string pstrSymbol ,
            int pintCurrRow ,
            int pintAbsLastCol )
        {
#if SEND_JSON_TO_FILE
            const string DEBUG_FILE_NAME_TEMPLATE = @"F:\Source_Code\Visual_Studio\Projects\_Laboratory\StockTickerSparklines\NOTES\strResponse_{0}_{1}.TXT";
#endif  // #if SEND_JSON_TO_FILE

            StockTickerEngine tickerEngine = StockTickerEngine.GetTheSingleInstance ( );

            string strQueryString = StockTickerEngine.BuildQueryString (
                StockTickerEngine.ApiFunction.TIME_SERIES_DAILY_ADJUSTED ,
                pstrSymbol );
            RestClient _restClient = new RestClient (
                Properties.Settings.Default.apiEndpoint ,
                RestClient.HttpVerb.Get ,
                RestClient.ContentType.JSON );
            string strResponse = _restClient.MakeRequest ( strQueryString );

#if SEND_JSON_TO_TRACE
            TraceLogger.WriteWithBothTimesLabeledLocalFirst (
                string.Format (
                    @"JSON response returned by API:{1}{1}{0}" ,
                    strResponse ,
                    Environment.NewLine ) );
#endif  // #if SEND_JSON_TO_TRACE

            if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
            {
                RestClient.ErrorResponse error = Newtonsoft.Json.JsonConvert.DeserializeObject<RestClient.ErrorResponse> (
                    strResponse.ApplyFixups (
                        new StringFixups.StringFixup [ ]
                        {
                            new StringFixups.StringFixup (
                                @"Error Message" ,          // Replace this ...
                                @"ErrorMessage" )           // ... with this
                        } ) );

                TraceLogger.WriteWithBothTimesLabeledLocalFirst (
                    string.Format (
                        @"The stock ticker API returned the following error message: {0}" ,
                        error.ErrorMessage ) );             // Format Item 0: the following error message: {0}
                txtMessage.Text = error.ErrorMessage;

                return null;
            }   // TRUE (unanticipated outcome) block, if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
            else
            {
                TraceLogger.WriteWithBothTimesLabeledLocalFirst ( @"The response looks good." );
                StringFixups.StringFixup [ ] stringFixups = StockTickerEngine.LoadStringFixups (
                    @"TIME_SERIES_DAILY_ResponseMap" );
                StringFixups responseStringFixups = new StringFixups ( stringFixups );
                string strFixedUp_Pass_1 = responseStringFixups.ApplyFixups ( strResponse );
                string strFixedUp_Pass_2 = ApplyFixups_Pass_2 ( strFixedUp_Pass_1 );

#if SEND_JSON_TO_FILE
                System.IO.File.WriteAllText (
                    string.Format (
                        DEBUG_FILE_NAME_TEMPLATE ,
                        pstrSymbol ,
                        @"Raw" ) ,
                    strResponse );
                System.IO.File.WriteAllText (
                    string.Format (
                        DEBUG_FILE_NAME_TEMPLATE ,
                        pstrSymbol ,
                        @"Pass_1" ) ,
                    strFixedUp_Pass_1 );
                System.IO.File.WriteAllText (
                    string.Format (
                        DEBUG_FILE_NAME_TEMPLATE ,
                        pstrSymbol ,
                        @"Pass_2" ) ,
                    strFixedUp_Pass_2 );
#endif  // #IF SEND_JSON_TO_FILE

                DailyTimeSeriesResponse timeSeriesResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<DailyTimeSeriesResponse> (
                    strFixedUp_Pass_2 );
                txtMessage.Text = Properties.Resources.MSG_HAVE_HISTORY;

                return timeSeriesResponse;
            }   // FALSE (anticipated outcome) block, if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
        }   // private DailyTimeSeriesResponse GetHistoryForSymbol


        private int FixNextItem (
            StringBuilder pbuilder ,
            int pintLastMatch )
        {   // This method references private instance member __fIsFirstPass several times.
            const string FIRST_ITEM_BREAK_ANTE = "[\n        \"";               // Ante: },\n        "
            const string SUBSEQUENT_ITEM_BREAK_ANTE = "},\n        \"";         // Ante: },\n        "

            string strInput = pbuilder.ToString ( );
            int intMatchPosition = strInput.IndexOf (
                __fIsFirstPass
                    ? FIRST_ITEM_BREAK_ANTE
                    : SUBSEQUENT_ITEM_BREAK_ANTE ,
                pintLastMatch );

            if ( intMatchPosition > ListInfo.INDEXOF_NOT_FOUND )
            {
                return FixThisItem (
                    strInput ,
                    intMatchPosition ,
                    __fIsFirstPass
                        ? FIRST_ITEM_BREAK_ANTE.Length
                        : SUBSEQUENT_ITEM_BREAK_ANTE.Length ,
                    pbuilder );
            }   // TRUE (At least one match remains.) block, if ( intMatchPosition > ListInfo.INDEXOF_NOT_FOUND )
            else
            {
                return ListInfo.INDEXOF_NOT_FOUND;
            }   // FALSE (All matches have been found.) block, if ( intMatchPosition > ListInfo.INDEXOF_NOT_FOUND )
        }   // private int FixNextItem


        private int FixThisItem (
            string pstrInput ,
            int pintMatchPosition ,
            int pintMatchLength ,
            StringBuilder psbOut )
        {
            const string FIRST_ITEM_BREAK_POST = "\n        {\n            \"Activity_Date\": \"";        // Post: },\n        {\n        {\n        "Activity_Date": "
            const string SUBSEQUENT_ITEM_BREAK_POST = ",\n        {\n            \"Activity_Date\": \"";    // Post: },\n        {\n        {\n        "Activity_Date": "

            const int DATE_TOKEN_LENGTH = 11;
            const int DATE_TOKEN_SKIP_CHARS = DATE_TOKEN_LENGTH + 3;

            int intSkipOverMatchedCharacters = pintMatchPosition + pintMatchLength;

            psbOut.Clear ( );

            psbOut.Append ( pstrInput.Substring (
                ListInfo.SUBSTR_BEGINNING ,
                ArrayInfo.OrdinalFromIndex ( pintMatchPosition ) ) );
            psbOut.Append ( __fIsFirstPass
                ? FIRST_ITEM_BREAK_POST
                : SUBSEQUENT_ITEM_BREAK_POST );
            psbOut.Append ( pstrInput.Substring (
                intSkipOverMatchedCharacters ,
                DATE_TOKEN_LENGTH ) );
            psbOut.Append ( SpecialCharacters.COMMA );
            psbOut.Append ( pstrInput.Substring ( intSkipOverMatchedCharacters + DATE_TOKEN_SKIP_CHARS ) );

            int rintSearchResumePosition =   pintMatchPosition 
                                           + ( __fIsFirstPass 
                                                    ? FIRST_ITEM_BREAK_POST.Length
                                                    : SUBSEQUENT_ITEM_BREAK_POST.Length );
            __fIsFirstPass = false;     // Putting this here allows execution to be unconditional.

            return ArrayInfo.OrdinalFromIndex ( rintSearchResumePosition );
        }   // private int FixThisItem


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
        }   // private void PopulateRowFromSearchResult method


        private void PruneSymbolsList ( )
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
                            keptRows [ intRowIndex ].RowIndex ,                 // int pintSourceRowIndex
                            ArrayInfo.OrdinalFromIndex ( intRowIndex ) ,        // int pintDestinationRowIndex
                            __intAbsLastCol ,                                   // int pintAbsLastCol
                            CellMovementDirection.Up ,
                            xlWork.ActiveSheet.Cells );                         // Cells pwksCells
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
        }   // private void PruneSymbolsList method


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
        }   // private void ReportException


        private void ResetForm ( )
        {
            __enmProcessingState = ProcessingState.ResettingForm;
            txtSearchString.Text = SpecialStrings.EMPTY_STRING;

            xlWork.ActiveSheet.Clear (
                ArrayInfo.ARRAY_FIRST_ELEMENT ,
                ArrayInfo.ARRAY_FIRST_ELEMENT ,
                __intAbsLastRow ,
                __intAbsLastCol );

            __intAbsLastRow = ArrayInfo.ARRAY_INVALID_INDEX;
            __intAbsLastCol = ArrayInfo.ARRAY_INVALID_INDEX;

            cmdExportToExcel.IsEnabled = false;
            cmdGetHistory.IsEnabled = false;
            cmdPruneSelections.IsEnabled = false;
            cmdResetForm.IsEnabled = false;

            txtMessage.Text = Properties.Resources.MSG_ENTER_SEARCH_STRING;

            xlWork.Invalidate ( );
            xlWork.UpdateLayout ( );

            txtSearchString.Focus ( );
            __enmProcessingState = ProcessingState.Idle;
        }   // private void ResetForm


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

            DataValidator validChoices = DataValidator.CreateListValidator (
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

            AutoSetColumnWidths (
                xlWork ,
                __intAbsLastCol );

            this.xlWork.ActiveSheet.Protect = true;
        }   // private void ShowSearchResults
        #endregion  // Private Instance Worker Methods


        #region Private Static Methods
        private static void AutoSetColumnWidths (
            GrapeCity.Windows.SpreadSheet.UI.GcSpreadSheet pwsWorkSheet ,
            int pintAbsoluteLastColumn )
        {
            //  --------------------------------------------------------
            //  Auto-fit the column widths.
            //  --------------------------------------------------------

            for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intColIndex <= pintAbsoluteLastColumn ;
                      intColIndex++ )
            {
                pwsWorkSheet.View.AutoFitColumn ( intColIndex );
            }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex <= pintAbsoluteLastColumn ; intColIndex++ )
        }   // private static void AutoSetColumnWidths


        private static void CreateSparklines (
            int pintFirstDetailColumn ,
            int pintDetailColumnCount ,
            int pintSparklineColumnIndex ,
            int [ ] paintSparklineRowAssignments ,
            GrapeCity.Windows.SpreadSheet.UI.GcSpreadSheet pxlWork )
        {
            for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intJ < paintSparklineRowAssignments.Length ;
                      intJ++ )
            {
                int intSparklineRow = paintSparklineRowAssignments [ intJ ];
                CellRange range = new CellRange (
                    intSparklineRow ,
                    pintFirstDetailColumn ,
                    MagicNumbers.PLUS_ONE ,
                    pintDetailColumnCount );
                SparklineSetting setting = new SparklineSetting ( );
                setting.AxisColor = SystemColors.ActiveBorderColor;
                setting.LineWeight = 1;
                setting.ShowMarkers = true;
                setting.MarkersColor = Color.FromRgb ( 255 , 0 , 128 );

                setting.ShowFirst = true;
                setting.ShowLow = true;
                setting.ShowLast = true;

                setting.HighMarkerColor = Color.FromRgb ( 49 , 78 , 111 );
                setting.LastMarkerColor = Color.FromRgb ( 0 , 255 , 255 );
                setting.LowMarkerColor = Color.FromRgb ( 255 , 255 , 0 );
                setting.NegativeColor = Color.FromRgb ( 255 , 255 , 0 );
                pxlWork.ActiveSheet.SetSparkline (
                    intSparklineRow ,                       // int              row             = Row of the cell into which to insert the graph
                    pintSparklineColumnIndex ,              // int              column          = Column of the cell into which to insert the graph
                    range ,                                 // CellRange        dataRange       = Data range from which to draw the graph
                    DataOrientation.Horizontal ,            // DataOrientation  dataOrientation = The data orientation
                    SparklineType.Line ,                    // SparklineType    type            = The sparkline type to draw
                    setting );                              // SparklineSetting setting         = The sparkline settings (colors and such)
                pxlWork.Invalidate ( );                     // Force the worksheet to repaint itself.
            }   // for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ; intJ < paintSparklineRowAssignments.Length ; intJ++ )
        }   // private static void CreateSparklines


        private static int [ ] MakeRoomForHistory (
            Worksheet pwsActiveWorkSheet ,
            int pintAbsLastRow ,
            int pintAbsLastCol ,
            int pintAdditionalRowsNeeded )
        {
            int [ ] raintIndexOfIssueRows = new int [ pintAbsLastRow ];

            int intOldIssueRow = ArrayInfo.ARRAY_SECOND_ELEMENT;

            for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intJ < raintIndexOfIssueRows.Length ;
                      intJ++ )
            {
                if ( intJ == ArrayInfo.ARRAY_FIRST_ELEMENT )
                {
                    raintIndexOfIssueRows [ intJ ] = intOldIssueRow;
                }   // TRUE (The first line stays put.) block, if ( intJ == ArrayInfo.ARRAY_FIRST_ELEMENT )
                else
                {
                    raintIndexOfIssueRows [ intJ ] = intOldIssueRow + ( pintAdditionalRowsNeeded * ArrayInfo.IndexFromOrdinal ( intOldIssueRow ) );
                }   // FALSE (Subsequent lines move down.) block, if ( intJ == ArrayInfo.ARRAY_FIRST_ELEMENT )

                intOldIssueRow++;
            }   // for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ; intJ < raintIndexOfIssueRows.Length ; intJ++ )

            //  ---------------------------------------------------------------
            //  Row movement must be from the bottom up.
            //  ---------------------------------------------------------------

            for ( int intCurrRow = pintAbsLastRow ;
                      intCurrRow > ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intCurrRow-- )
            {
                MovePopulatedRow (
                    intCurrRow ,                                                // int pintSourceRowIndex
                    raintIndexOfIssueRows [ ArrayInfo.IndexFromOrdinal (        // int pintDestinationRowIndex
                        intCurrRow ) ] ,                                        // The array of new locations is one subscript behind the array of cells.
                    pintAbsLastCol ,                                            // int pintAbsLastCol
                    CellMovementDirection.Down ,
                    pwsActiveWorkSheet.Cells );                                 // Cells pwksCells
            }   // for ( int intCurrRow = pintAbsLastRow ; intCurrRow > ArrayInfo.ARRAY_FIRST_ELEMENT ; intCurrRow-- )

            return raintIndexOfIssueRows;
        }   // private static int [ ] MakeRoomForHistory


        private static void MovePopulatedRow (
            int pintSourceRowIndex ,
            int pintDestinationRowIndex ,
            int pintAbsLastCol ,
            CellMovementDirection penmCellMovementDirection ,
            Cells pwksCells )
        {
            if ( MoveIsSafe ( pintSourceRowIndex , pintDestinationRowIndex , penmCellMovementDirection ) )
            {
                for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                          intColIndex <= pintAbsLastCol ;
                          intColIndex++ )
                {
                    pwksCells [ pintDestinationRowIndex , intColIndex ].Value = pwksCells [ pintSourceRowIndex , intColIndex ].Value;
                    pwksCells [ pintSourceRowIndex , intColIndex ].Value = null;
                }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex <= pintAbsLastCol ; intColIndex++ )
            }   // if ( MoveIsSafe ( pintSourceRowIndex , pintDestinationRowIndex , penmCellMovementDirection ) )
        }   // private static void MovePopulatedRow


        private static bool MoveIsSafe (
            int pintSourceRowIndex ,
            int pintDestinationRowIndex ,
            CellMovementDirection penmCellMovementDirection )
        {
            if ( penmCellMovementDirection == CellMovementDirection.Up )
            {
                return ( pintSourceRowIndex > pintDestinationRowIndex );
            }   // TRUE (Rows are being moved UP to lower-numbered rows, closer to the top of the sheet.) block, if ( penmCellMovementDirection == CellMovementDirection.Up )
            else
            {
                return ( pintDestinationRowIndex > pintSourceRowIndex );
            }   // FALSE (Rows are being moved DOWN to higher-numbered rows, further from the top of the sheet.) block, if ( penmCellMovementDirection == CellMovementDirection.Up )
        }   // private static bool MoveIsSafe


        private static SparklineRowAssignments StoreTimeSeriesInWorksheet (
            int pintOriginRow ,
            int pintOriginColumn ,
            string [ ] pastrRowLabels ,
            Time_Series_Daily [ ] patsdTimeSeriesDaily ,
            Worksheet pwsActiveSheet )
        {
            int intFirstDataColumn = pintOriginColumn + ArrayInfo.NEXT_INDEX;

            //  ----------------------------------------------------------------
            //  Fill the label column.
            //  ----------------------------------------------------------------

            for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intJ < pastrRowLabels.Length ;
                      intJ++ )
            {
                pwsActiveSheet.Cells [ pintOriginRow + intJ , pintOriginColumn ].Value = pastrRowLabels [ intJ ];
            }   // for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ; intJ < pastrRowLabels.Length ; intJ++ )

            //  ----------------------------------------------------------------
            //  Before the next part can succeed, the number of columns in the
            //  worksheet must be increased from its default initial value of
            //  one hundred. The increase must account for the first data column
            //  plus one, to allow for the sparkline.
            //  ----------------------------------------------------------------

            pwsActiveSheet.ColumnCount = intFirstDataColumn + patsdTimeSeriesDaily.Length + ArrayInfo.NEXT_INDEX;

            //  ----------------------------------------------------------------
            //  Beginning with the next column, fill down with field values from
            //  an element of the Time_Series_Daily array, working across the
            //  sheet until all Time_Series_Daily array elments have been used.
            //  ----------------------------------------------------------------

            SparklineRowAssignments assignments = new SparklineRowAssignments (
                intFirstDataColumn ,
                patsdTimeSeriesDaily.Length );

            for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intJ < patsdTimeSeriesDaily.Length ;
                      intJ++ )
            {
                StoreTimeSeriesItemInWorksheet (
                    patsdTimeSeriesDaily [ intJ ] ,         // Time_Series_Daily        ptsdTimeSeriesDailyItem
                    pintOriginRow ,                         // int                      pintOriginRow
                    intFirstDataColumn + intJ ,             // int                      pintDestinationColumn
                    pwsActiveSheet ,                        // Worksheet                pwsActiveSheet
                    assignments );                          // SparklineRowAssignments  passignments
            }   // for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ; intJ < patsdTimeSeriesDaily.Length ; intJ++ )

            return assignments;
        }   // private static int StoreTimeSeriesInWorksheet


        private static void StoreTimeSeriesItemInWorksheet (
            Time_Series_Daily ptsdTimeSeriesDailyItem ,
            int pintOriginRow ,
            int pintDestinationColumn ,
            Worksheet pwsActiveSheet ,
            SparklineRowAssignments passignments )
        {
            int intCurrRow = pintOriginRow;

            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.Activity_Date );

            passignments.Open = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.Open );

            passignments.High = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.High );

            passignments.Low = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.Low );

            passignments.Close = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.Close );

            passignments.AdjustedClose = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.AdjustedClose );

            passignments.Adjustment = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = ptsdTimeSeriesDailyItem.GetAdjustment ( );

            passignments.Volume = intCurrRow;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = Time_Series_Daily.ConvertToAppropriateType ( ptsdTimeSeriesDailyItem.Volume );

            //  ----------------------------------------------------------------
            //  The last two change too infrequently to warrant a graph, and the
            //  final increment paves the way should another row be added.
            //  ----------------------------------------------------------------

            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = ptsdTimeSeriesDailyItem.DividendAmount;
            pwsActiveSheet.Cells [ intCurrRow++ , pintDestinationColumn ].Value = ptsdTimeSeriesDailyItem.SplitCoefficient;
        }   // private static void StoreTimeSeriesItemInWorksheet


        private static int UpdateLastRow ( string [ ] pastrTimeSeriesLabels , int pintCurrRow , int pintCurrentLastRow )
        {
            int intLastRowCandidate = pintCurrRow + pastrTimeSeriesLabels.Length;

            if ( pintCurrentLastRow < intLastRowCandidate )
            {
                return intLastRowCandidate;
            }   // TRUE (Last row is out of date with respect to the worksheet's content.) block, if ( pintCurrentLastRow < intLastRowCandidate )
            else
            {
                return pintCurrentLastRow;
            }   // FALSE (Last row is current with respect to the worksheet's content.) block, if ( pintCurrentLastRow < intLastRowCandidate )
        }   // private static int UpdateLastRow
        #endregion  // Private Static Methods


        #region Private Custom Instance Storage (Double underscores differentiate these from privates inherited from the base class.)
        private bool __fIsFirstPass = true;

        private int __intAbsLastRow = ArrayInfo.ARRAY_INVALID_INDEX;
        private int __intAbsLastCol = ArrayInfo.ARRAY_INVALID_INDEX;

        private ProcessingState __enmProcessingState = ProcessingState.Idle;
        #endregion  // Private Custom Instance Storage


        #region Nested Private Classes
        private class KeptRow : IComparable<KeptRow>
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
#endregion  // Nested Private Classes
    }   // public partial class MainWindow
}   // partial namespace StockTickerSparklines