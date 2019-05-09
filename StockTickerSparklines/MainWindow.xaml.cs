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
        #region Constructor
        public MainWindow ( )
        {
            InitializeComponent ( );
            this.Title = System.Reflection.Assembly.GetExecutingAssembly ( ).GetName ( ).Name;
        }   // public MainWindow default constructor
        #endregion  // Constructor


        #region Event Delegates
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

                    this.txtMessage.Text = Properties.Resources.MSG_SYMBOLS_FOUND;
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
            PruneSymbolsList ( );
        }   // CmdPruneSelections_Click event delegate


        private void CmdGetHistory_Click ( object sender , RoutedEventArgs e )
        {
            GetHistoryForSelectedSymbols ( );
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
        #endregion  // Event Delegates


        #region Private Worker Methods
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
            const int SYMBOL_COLUMN = 1;

            try
            {
                for ( int intCurrRow = ArrayInfo.ARRAY_SECOND_ELEMENT ;
                          intCurrRow <= __intAbsLastRow ;
                          intCurrRow++ )
                {
                    GetHistoryForSymbol (
                        xlWork.ActiveSheet.Cells [ intCurrRow , SYMBOL_COLUMN ].Text ,
                        intCurrRow ,
                        __intAbsLastCol );
                }   // for ( int intCurrRow = ArrayInfo.ARRAY_SECOND_ELEMENT ; intCurrRow <= __intAbsLastRow ; intCurrRow++ )
            }
            catch ( Exception ex )
            {
                ReportException ( ex );
            }   // catch ( Exception ex )
        }   // private void GetHistoryForSelectedSymbols


        private void GetHistoryForSymbol (
            string pstrSymbol ,
            int pintCurrRow ,
            int pintAbsLastCol )
        {
            const string DEBUG_FILE_NAME_TEMPLATE = @"F:\Source_Code\Visual_Studio\Projects\_Laboratory\StockTickerSparklines\NOTES\strResponse_{0}_{1}.TXT";

            StockTickerEngine tickerEngine = StockTickerEngine.GetTheSingleInstance ( );

            string strQueryString = StockTickerEngine.BuildQueryString (
                StockTickerEngine.ApiFunction.TIME_SERIES_DAILY_ADJUSTED ,
                pstrSymbol );
            RestClient _restClient = new RestClient (
                Properties.Settings.Default.apiEndpoint ,
                RestClient.HttpVerb.Get ,
                RestClient.ContentType.JSON );
            string strResponse = _restClient.MakeRequest ( strQueryString );

            TraceLogger.WriteWithBothTimesLabeledLocalFirst (
                string.Format (
                    @"JSON response returned by API:{1}{1}{0}" ,
                    strResponse ,
                    Environment.NewLine ) );

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
            }   // TRUE (unanticipated outcome) block, if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
            else
            {
                TraceLogger.WriteWithBothTimesLabeledLocalFirst ( @"The response looks good." );
                StringFixups.StringFixup [ ] stringFixups = StockTickerEngine.LoadStringFixups (
                    @"TIME_SERIES_DAILY_ResponseMap" );
                StringFixups responseStringFixups = new StringFixups ( stringFixups );
                string strFixedUp_Pass_1 = responseStringFixups.ApplyFixups ( strResponse );
                string strFixedUp_Pass_2 = ApplyFixups_Pass_2 ( strFixedUp_Pass_1 );

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

                //DailyTimeSeriesResponse timeSeriesResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<DailyTimeSeriesResponse> (
                //    responseStringFixups.ApplyFixups (
                //        strResponse ) );
                txtMessage.Text = Properties.Resources.MSG_HAVE_HISTORY;
            }   // FALSE (anticipated outcome) block, if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
        }   // private void GetHistoryForSymbol


        private string ApplyFixups_Pass_2 ( string pstrFixedUp_Pass_1 )
        {
            const string TSD_LABEL_ANTE = "\"TimeSeriesDaily\": {";                                     // Ante: "TimeSeriesDaily": {
            const string TSD_LABEL_POST = "\"Time_Series_Daily\" : [";                                  // Post: "Time_Series_Daily": [

            StringBuilder builder1 = new StringBuilder ( pstrFixedUp_Pass_1.Length * MagicNumbers.PLUS_TWO );

            builder1.Append (
                pstrFixedUp_Pass_1.Replace (
                    TSD_LABEL_ANTE ,
                    TSD_LABEL_POST ) );

            int intLastMatch = builder1.Length + ArrayInfo.NEXT_INDEX;

            while ( intLastMatch > ListInfo.INDEXOF_NOT_FOUND )
            {
                intLastMatch = FixNextItem (
                    builder1 ,
                    intLastMatch );
            }   // while ( intLastMatch > ListInfo.INDEXOF_NOT_FOUND )

            return builder1.ToString ( );
        }   // private string ApplyFixups_Pass_2


        private int FixNextItem (
            StringBuilder pbuilder ,
            int pintLastMatch )
        {
            const string ITEM_BREAK_ANTE = "},\n        \"";                                            // Ante: },\n        "

            string strInput = pbuilder.ToString ( );

            int intMatchPosition = strInput.IndexOf (
                ITEM_BREAK_ANTE ,
                pintLastMatch );

            if ( intMatchPosition > ListInfo.INDEXOF_NOT_FOUND )
            {
                return FixThisItem (
                    strInput ,
                    intMatchPosition ,
                    ITEM_BREAK_ANTE.Length ,
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
            const string ITEM_BREAK_POST = "},\n        {\n        {\n        \"Activity_Date\": \"";   // Post: },\n        {\n        {\n        "Activity_Date": "

            const int DATE_TOKEN_LENGTH = 11;
            const int DATE_TOKEN_SKIP_CHARS = DATE_TOKEN_LENGTH + 4;

            int intSkipOverMatchedCharacters = pintMatchPosition + pintMatchLength;

            psbOut.Clear ( );

            psbOut.Append ( pstrInput.Substring (
                ListInfo.SUBSTR_BEGINNING ,
                ArrayInfo.OrdinalFromIndex ( pintMatchPosition ) ) );
            psbOut.Append ( ITEM_BREAK_POST );
            psbOut.Append ( pstrInput.Substring (
                intSkipOverMatchedCharacters ,
                DATE_TOKEN_LENGTH ) );
            psbOut.Append ( SpecialCharacters.COMMA );
            psbOut.Append ( pstrInput.Substring ( intSkipOverMatchedCharacters + DATE_TOKEN_SKIP_CHARS ) );

            return ArrayInfo.OrdinalFromIndex ( pintMatchPosition + ITEM_BREAK_POST.Length );
        }   // private int FixThisItem


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
        #endregion  // Private Worker Methods


        #region Private Custom Instance Storage (Double underscores differentiate these from privates inherited from the base class.)
        private int __intAbsLastRow = ArrayInfo.ARRAY_INVALID_INDEX;
        private int __intAbsLastCol = ArrayInfo.ARRAY_INVALID_INDEX;
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