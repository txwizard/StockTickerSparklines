using System;
using System.Text;

using WizardWrx;

namespace StockTickerSparklines
{
    class SymbolInfo : IComparable<SymbolInfo>
    {
        private SymbolInfo ( )
        {
        }   // SymbolInfo constructor (1 of 2) is marked private to force instances to be initalized.

        public SymbolInfo ( string pstrOutputValue )
        {
            StringBuilder labelBuilder = new StringBuilder ( pstrOutputValue.Length );
            char [ ] achrValueCharacters = pstrOutputValue.ToCharArray ( );
            int intTempOrder = ListInfo.LIST_IS_EMPTY;

            for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intJ < achrValueCharacters.Length ;
                      intJ++ )
            {
                switch ( achrValueCharacters [ intJ ] )
                {
                    case SpecialCharacters.UNDERSCORE_CHAR:
                        break;  // Skip it.
                    case SpecialCharacters.CHAR_NUMERAL_0:
                        break;  // Skip it.
                    case SpecialCharacters.CHAR_NUMERAL_1:
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + MagicNumbers.PLUS_ONE;
                        break;
                    case SpecialCharacters.CHAR_NUMERAL_2:
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + MagicNumbers.PLUS_TWO;
                        break;
                    case '3':
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + 3;
                        break;
                    case '4':
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + 4;
                        break;
                    case '5':
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + 5;
                        break;
                    case '6':
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + 6;
                        break;
                    case SpecialCharacters.CHAR_NUMERAL_7:
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + MagicNumbers.PLUS_SEVEN;
                        break;
                    case '8':
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + 8;
                        break;
                    case '9':
                        intTempOrder = ( MagicNumbers.EXACTLY_TEN * intTempOrder ) + 9;
                        break;
                    default:
                        labelBuilder.Append ( 
                            ( labelBuilder.Length == ListInfo.LIST_IS_EMPTY )
                            ? ToUpper ( achrValueCharacters [ intJ ] ) 
                            : achrValueCharacters [ intJ ] );
                        break;
                }   // switch ( achrValueCharacters [ intJ ] )
            }   // for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ; intJ < achrValueCharacters.Length ; intJ++ )

            Order = intTempOrder;
            Label = labelBuilder.ToString ( );
        }   // SymbolInfo constructor (2 of 2) is marked public, and its argument is enough to initialize all properties.

        public int Order
        {
            get; private set;
        }   // Order property is read-only to the outside world.

        public string Label
        {
            get; private set;
        }   // Label property is read-only to the outside world.

        public override bool Equals ( object obj )
        {
            SymbolInfo other = obj as SymbolInfo;
            return Order.Equals ( other.Order );
        }   // Equals operator overload

        public override int GetHashCode ( )
        {
            return Order.GetHashCode ( );
        }   // GetHashCode overload

        public override string ToString ( )
        {
            return string.Format (
                Properties.Resources.SYMBOLINFO_TOSTRING_TEMPLATE ,             // Format Control String lives in the language-neutral string resource table.
                Order ,                                                         // Format Item 0: Order = {0},
                Label );                                                        // Format Item 1: Label = {1}
        }   // ToString overload

        int IComparable<SymbolInfo>.CompareTo ( SymbolInfo other )
        {
            return Order.CompareTo ( other.Order );
        }   // CompareTo method

        private char ToUpper ( char pchr )
        {   // ToDo: There is a mathematical way to do this that's much more efficient.
            string strOfOneChar = pchr.ToString ( );
            string strUpperChar = strOfOneChar.ToUpper ( );
            char [ ] rachrUpperChar = strUpperChar.ToCharArray ( );

            return rachrUpperChar [ ArrayInfo.ARRAY_FIRST_ELEMENT ];
        }   // ToUpper method
    }   // class SymbolInfo
}   // namespace StockTickerSparklines