using System;

namespace StockTickerSparklines
{
    class TickerSymbol : IComparable<TickerSymbol>
    {
        #region Public Properties
        /// <summary>
        /// Symbol is the ticker symbol, the key to getting stock quotes.
        /// </summary>
        public string Symbol { get; set; }

        /// <summary>
        /// The Issue Name is the name of the issue; for stocks, this is usually
        /// the name of the issuing company.
        /// </summary>
        public string IssueName { get; set; }

        /// <summary>
        /// This flag is used internally to designate issues to be added to the
        /// list of stocks for which to get daily closing proces from which to
        /// generate the sparkline graphs.
        /// </summary>
        public bool IsSelected { get; set; }
        #endregion  // Public Properties


        #region IComparable Interface Implementation
        int IComparable<TickerSymbol>.CompareTo ( TickerSymbol that )
        {
            return Symbol.CompareTo ( that.Symbol );
        }   // int IComparable<TickerSymbol>.CompareTo
        #endregion  // IComparable Interface Implementation


        #region Overrides of Methods on Base Class, Object
        public override bool Equals ( object obj )
        {
            TickerSymbol symbol = obj as TickerSymbol;
            return Symbol.Equals ( symbol.Symbol );
        }   // public override bool Equals


        public override int GetHashCode ( )
        {
            return Symbol.GetHashCode ( );
        }   // public override int GetHashCode


        public override string ToString ( )
        {
            return string.Format (
                Properties.Resources.TPL_TICKER_SYMBOL_TOSTRING ,               // Format control string lives in a managed string resource for easier internationalization.
                Symbol ,                                                        // Format Item 0: Symbol = {0}
                IssueName ,                                                     // Format Item 1: Issuer Name = {1}
                IsSelected );                                                   // Format Item 2: Selected {2}
        }   // public override string ToString
        #endregion  // Overrides of Methods on Base Class, Object
    }   // class TickerSymbol is implicitly internal.
}   // partial namespace StockTickerSparklines