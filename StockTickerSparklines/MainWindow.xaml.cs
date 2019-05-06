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
            try
            {
                StockTickerEngine tickerEngine = StockTickerEngine.GetTheSingleInstance ( );
                TickerSymbolsCollection symbols = tickerEngine.Search ( txtSearchString.Text );

                if ( symbols != null )
                {
                    this.txtMessage.Text = @"Symbols for you I have!";
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