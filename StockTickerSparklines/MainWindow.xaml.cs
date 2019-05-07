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
            this.txtSearchString.Focus ( );
        }   // public MainWindow default constructor
 
 
        private void Button_Click ( object sender , RoutedEventArgs e )
        {
            MessageBox.Show (
                @"Click!" ,
                this.Title );
        }   // private void Button_Click event delegate

 
        private void TxtSearchString_TextChanged ( object sender , TextChangedEventArgs e )
        {
            TextBox textBox = sender as TextBox;

            this.cmdSearch.IsEnabled = ( textBox.Text.Length > ListInfo.EMPTY_STRING_LENGTH );
        }   // private void TxtSearchString_TextChanged event delegate
 
 
        private void CmdSearch_Click ( object sender , RoutedEventArgs e )
        {
            const int COL_IDX_SYMBOL = 0;
            const int COL_IDX_NAME = 1;
            const int COL_IDX_TYPE = 2;
            const int COL_IDX_REGION = 3;
            const int COL_IDX_MARKETOPEN = 4;
            const int COL_IDX_MARKETCLOSE = 5;
            const int COL_IDX_TIMEZONE = 6;
            const int COL_IDX_CURRENCY = 7;
            const int COL_IDX_MATCHSCORE = 8;

            try
            {
                StockTickerEngine tickerEngine = StockTickerEngine.GetTheSingleInstance ( );
                TickerSymbolMatches symbols = tickerEngine.Search ( txtSearchString.Text );

                if ( symbols != null )
                {
                    SymbolInfo [ ] symbolInfos = tickerEngine.GetSymbolInfos ( );

                    //  --------------------------------------------------------
                    //  Label the first nine columns.
                    //  --------------------------------------------------------

                    for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                              intColIndex < symbolInfos.Length ;
                              intColIndex++ )
                    {
                        this.xlWork.ActiveSheet.Cells [ ArrayInfo.ARRAY_FIRST_ELEMENT , intColIndex ].Value = symbolInfos [ intColIndex ].Label;
                    }   // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex < symbolInfos.Length ; intColIndex++ )

                    //  --------------------------------------------------------
                    //  Polulate the detail rows.
                    //  --------------------------------------------------------

                    for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                              intRowIndex < symbols.bestMatches.Length ;
                              intRowIndex++ )
                    {
                        for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                              intColIndex < symbolInfos.Length ;
                              intColIndex++ )
                        {
                            string strCellValue = null;

                            //  ------------------------------------------------
                            //  The column index determines which of the nine
                            //  values in the BestMatches array goes into the
                            //  current column. As a matter of habit, I always
                            //  code for the default case, even when it should
                            //  never happen.
                            //  ------------------------------------------------

                            switch ( intColIndex )
                            {
                                case COL_IDX_SYMBOL:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._1symbol;
                                    break;
                                case COL_IDX_NAME:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._2name;
                                    break;
                                case COL_IDX_TYPE:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._3type;
                                    break;
                                case COL_IDX_REGION:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._4region;
                                    break;
                                case COL_IDX_MARKETOPEN:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._5marketOpen;
                                    break;
                                case COL_IDX_MARKETCLOSE:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._6marketClose;
                                    break;
                                case COL_IDX_TIMEZONE:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._7timezone;
                                    break;
                                case COL_IDX_CURRENCY:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._8currency;
                                    break;
                                case COL_IDX_MATCHSCORE:
                                    strCellValue = symbols.bestMatches [ intRowIndex ]._9matchScore;
                                    break;
                                default:
                                    throw new InvalidOperationException ( string.Format (
                                        Properties.Resources.TPL_INTERNAL_ERROR_001 ,
                                        intColIndex ,                           // Format Item 0: intColIndex has an invalid value of {0}.
                                        symbolInfos.Length ) );                 // Format Item 1: Its value must always be less than {1}.
                            }   // switch ( intColIndex )

                            this.xlWork.ActiveSheet.Cells [ ArrayInfo.OrdinalFromIndex ( intRowIndex ) , intColIndex ].Value = strCellValue;
                        }   // // for ( int intColIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intColIndex < symbolInfos.Length ; intColIndex++ )
                    }   // for ( int intRowIndex = ArrayInfo.ARRAY_FIRST_ELEMENT ; intRowIndex < symbols.bestMatches.Length ; intRowIndex++ )

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
                    Title );
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
    }   // public partial class MainWindow
}   // partial namespace StockTickerSparklines