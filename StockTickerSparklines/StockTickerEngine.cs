using System;
using System.Collections.Generic;
using System.Text;

using WizardWrx;
using WizardWrx.AnyCSV;
using WizardWrx.Core;
using WizardWrx.EmbeddedTextFile;

namespace StockTickerSparklines
{
    class StockTickerEngine : GenericSingletonBase<StockTickerEngine>
    {
        enum ApiFunction
        {
            GLOBAL_QUOTE,
            SYMBOL_SEARCH,
            TIME_SERIES_DAILY_ADJUSTED,
        };  // ApiFunction enumeration


        public string Message
        {
            get; private set;
        }   // public string Message property (read-only)


        /// <summary>
        /// The body of this method is temporary; the end goal is for the body
        /// of GetSymbols to replace it, and for the private GetSymbols method
        /// to go away.
        /// </summary>
        /// <param name="pstrQueryString">
        /// The query string is fed from the form to return a list of stocks to
        /// go into the list of stocks to follow.
        /// </param>
        /// <returns>
        /// If the method succeeds, the return value is a TickerSymbolsCollection
        /// that contains at least one TickerSymbol object.
        /// </returns>
        internal TickerSymbolMatches Search ( string pstrQueryString )
        {
            // ToDo: Replace this body with that of GetSymbols, and delete that method.
            TickerSymbolMatches symbols = GetSymbols ( pstrQueryString );

            if ( symbols != null )
            {
                return symbols;
            }   // TRUE (anticipated outcome) block, if ( symbols != null )
            else
            {
                return null;
            }   // FALSE (unanticipated outcome) block, if ( symbols != null )
        }   // internal TickerSymbolsCollection Search


        public SymbolInfo [ ] GetSymbolInfos ( )
        {
            return _symbolInfo;
        }   // public SymbolInfo [ ] GetSymbolInfos


        /// <summary>
        /// Get the list of ticker symbols that match a specified
        /// <paramref name="pstrQueryString"/>.
        /// </summary>
        /// <param name="pstrQueryString">
        /// The query string is fed from the form to return a list of stocks to
        /// go into the list of stocks to follow.
        /// </param>
        /// <returns>
        /// If the method succeeds, the return value is a TickerSymbolsCollection
        /// that contains at least one TickerSymbol object.
        /// </returns>
        private TickerSymbolMatches GetSymbols ( string pstrQueryString )
        {
            string strQueryString = BuildQueryString (
                ApiFunction.SYMBOL_SEARCH ,
                pstrQueryString );
            _restClient = _restClient ?? new RestClient (
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
                this.Message = error.ErrorMessage;
                return null;
            }   // TRUE (unanticipated outcome) block, if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
            else
            {
                TraceLogger.WriteWithBothTimesLabeledLocalFirst ( @"The response looks good." );

                StringFixups.StringFixup [ ] stringFixups = LoadStringFixups ( @"SYMBOL_SEARCH_ResponseMasp" );
                _responseStringFixups = _responseStringFixups ?? new StringFixups ( stringFixups );
                _symbolInfo = _symbolInfo ?? new SymbolInfo [ stringFixups.Length ];

                for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                          intJ < stringFixups.Length ;
                          intJ++ )
                {
                    _symbolInfo [ intJ ] = new SymbolInfo ( stringFixups [ intJ ].OutputValue );
                }   // for ( int intJ = ArrayInfo.ARRAY_FIRST_ELEMENT ; intJ < stringFixups.Length ; intJ++ )

                return Newtonsoft.Json.JsonConvert.DeserializeObject<TickerSymbolMatches> (
                    _responseStringFixups.ApplyFixups (
                        strResponse ) );
            }   // FALSE (anticipated outcome) block, if ( RestClient.ErrorResponse.ResponseIsErrorMessage ( strResponse ) )
        }   // private TickerSymbolsCollection GetSymbols method


        /// <summary>
        /// Assemble the query string required to make a request to the stock
        /// ticker service.
        /// </summary>
        /// <param name="penmApiFunction">
        /// Each supported function is mapped to a member of the ApiFunction
        /// enumeration, which does double duty, cast to an integer, as the
        /// subscript into the _map array that looks up the corresponding
        /// ParameterName string to pair with the <paramref name="pstrParameterValue"/>
        /// string.
        /// </param>
        /// <param name="pstrParameterValue">
        /// The value specified in this string is the value that is paired with
        /// the parameter named that is mapped to the specified
        /// <paramref name="penmApiFunction"/> value.
        /// </param>
        /// <returns>
        /// The return value is the query string to submit to the single API
        /// endpoint.
        /// </returns>
        private string BuildQueryString (
            ApiFunction penmApiFunction ,
            string pstrParameterValue )
        {
            // Template:    ?function=SYMBOL_SEARCH&keywords=BA&apikey=YH5RG5INKJN1HCXL

            _map = _map ?? LoadMap ( );

            RestClient.QueryStringBuilder queryStringBuilder = new RestClient.QueryStringBuilder ( );

            FunctionMapItem mapItem = _map [ ( int ) penmApiFunction ];
            queryStringBuilder.AddParameter (
                @"function" ,
                mapItem.function.ToString ( ) );
            queryStringBuilder.AddParameter (
                mapItem.parameter ,
                pstrParameterValue );
            queryStringBuilder.AddParameter (
                @"apikey" ,
                Properties.Settings.Default.apikey );

            return queryStringBuilder.GetQueryString ( );
        }   // private string BuildQueryString method


        private static FunctionMapItem [ ] LoadMap ( )
        {
            const string FILE_NAME = @"ApiParameterMap.txt";
            const string LABEL_ROW = @"ApiFunction	ParameterName";

            const int EXPECTED_FIELD_COUNT = 2;

            string [ ] astrAllMapItems = Readers.LoadTextFileFromEntryAssembly ( FILE_NAME );

            Parser parser = new Parser (
                CSVParseEngine.DelimiterChar.Tab ,
                CSVParseEngine.GuardChar.DoubleQuote ,
                CSVParseEngine.GuardDisposition.Strip );
            FunctionMapItem [ ] rFunctionMaps = new FunctionMapItem [ ArrayInfo.IndexFromOrdinal ( astrAllMapItems.Length ) ];

            for ( int intI = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intI < astrAllMapItems.Length ;
                      intI++ )
            {
                if ( intI == ArrayInfo.ARRAY_FIRST_ELEMENT )
                {
                    if ( astrAllMapItems [ intI ] != LABEL_ROW )
                    {
                        throw new Exception (
                            string.Format (
                                Properties.Resources.ERRMSG_CORRUPTED_EMBBEDDED_RESOURCE_LABEL ,
                                new string [ ]
                                {
                                    FILE_NAME ,                                 // Format Item 0: internal resource {0}
                                    LABEL_ROW ,                                 // Format Item 1: Expected value = {1}
                                    astrAllMapItems [ intI ] ,                  // Format Item 2: Actual value   = {2}
                                    Environment.NewLine                         // Format Item 3: Platform-specific newline
                                } ) );
                    }   // if ( astrAllMapItems[intI] != LABEL_ROW )
                }   // TRUE (label row sanity check 1 of 2) block, if ( intI == ArrayInfo.ARRAY_FIRST_ELEMENT )
                else
                {
                    string [ ] astrFields = parser.Parse ( astrAllMapItems [ intI ] );

                    if ( astrFields.Length == EXPECTED_FIELD_COUNT )
                    {
                        rFunctionMaps [ ArrayInfo.IndexFromOrdinal ( intI ) ] = new FunctionMapItem
                        {
                            function = ( ApiFunction ) Enum.Parse (
                            typeof ( ApiFunction ) ,
                            astrFields [ ArrayInfo.ARRAY_FIRST_ELEMENT ] ) ,
                            parameter = astrFields [ ArrayInfo.ARRAY_SECOND_ELEMENT ]
                        };
                    }   // TRUE (anticipated outcome) block, if ( astrFields.Length == EXPECTED_FIELD_COUNT )
                    else
                    {
                        throw new Exception (
                            string.Format (
                                Properties.Resources.ERRMSG_CORRUPTED_EMBEDDED_RESOURCE_DETAIL ,
                                new object [ ]
                                {
                                    intI ,                                      // Format Item 0: Detail record {0}
                                    FILE_NAME ,                                 // Format Item 1: internal resource {1}
                                    EXPECTED_FIELD_COUNT ,                      // Format Item 2: Expected field count = {2}
                                    astrFields.Length ,                         // Format Item 3: Actual field count   = {3}
                                    astrAllMapItems [ intI ] ,                  // Format Item 4: Actual record        = {4}
                                    Environment.NewLine                         // Format Item 5: Platform-specific newline
                                } ) );
                    }   // FALSE (unanticipated outcome) block, if ( astrFields.Length == EXPECTED_FIELD_COUNT )
                }   // FALSE (detail row) block, if ( intI == ArrayInfo.ARRAY_FIRST_ELEMENT )
            }   // for ( int intI = ArrayInfo.ARRAY_FIRST_ELEMENT ; intI < astrAllMapItems.Length ; intI++ )

            return rFunctionMaps;
        }   // private static FunctionMapItem LoadMap method


        private StringFixups.StringFixup [ ] LoadStringFixups ( string pstrEmbeddedResourceName )
        {
            const string LABEL_ROW = @"JSON	VS";
            const string TSV_EXTENSION = @".txt";

            const int STRING_PER_RESPONSE = ArrayInfo.ARRAY_FIRST_ELEMENT;
            const int STRING_FOR_JSONCONVERTER = STRING_PER_RESPONSE + ArrayInfo.NEXT_INDEX;
            const int EXPECTED_FIELD_COUNT = STRING_FOR_JSONCONVERTER + ArrayInfo.NEXT_INDEX;

            string strEmbeddResourceFileName = string.Concat (
                pstrEmbeddedResourceName ,
                TSV_EXTENSION );
            string [ ] astrAllMapItems = Readers.LoadTextFileFromEntryAssembly ( strEmbeddResourceFileName );
            Parser parser = new Parser (
                CSVParseEngine.DelimiterChar.Tab ,
                CSVParseEngine.GuardChar.DoubleQuote ,
                CSVParseEngine.GuardDisposition.Strip );
            StringFixups.StringFixup [ ] rFunctionMaps = new StringFixups.StringFixup [ ArrayInfo.IndexFromOrdinal ( astrAllMapItems.Length ) ];

            for ( int intI = ArrayInfo.ARRAY_FIRST_ELEMENT ;
                      intI < astrAllMapItems.Length ;
                      intI++ )
            {
                if ( intI == ArrayInfo.ARRAY_FIRST_ELEMENT )
                {
                    if ( astrAllMapItems [ intI ] != LABEL_ROW )
                    {
                        throw new Exception (
                            string.Format (
                                Properties.Resources.ERRMSG_CORRUPTED_EMBBEDDED_RESOURCE_LABEL ,
                                new string [ ]
                                {
                                    strEmbeddResourceFileName ,                 // Format Item 0: internal resource {0}
                                    LABEL_ROW ,                                 // Format Item 1: Expected value = {1}
                                    astrAllMapItems [ intI ] ,                  // Format Item 2: Actual value   = {2}
                                    Environment.NewLine                         // Format Item 3: Platform-specific newline
                                } ) );
                    }   // if ( astrAllMapItems[intI] != LABEL_ROW )
                }   // TRUE (label row sanity check 1 of 2) block, if ( intI == ArrayInfo.ARRAY_FIRST_ELEMENT )
                else
                {
                    string [ ] astrFields = parser.Parse ( astrAllMapItems [ intI ] );

                    if ( astrFields.Length == EXPECTED_FIELD_COUNT )
                    {
                        rFunctionMaps [ ArrayInfo.IndexFromOrdinal ( intI ) ] = new StringFixups.StringFixup (
                            astrFields [ STRING_PER_RESPONSE ] ,
                            astrFields [ STRING_FOR_JSONCONVERTER ] );
                    }   // TRUE (anticipated outcome) block, if ( astrFields.Length == EXPECTED_FIELD_COUNT )
                    else
                    {
                        throw new Exception (
                            string.Format (
                                Properties.Resources.ERRMSG_CORRUPTED_EMBEDDED_RESOURCE_DETAIL ,
                                new object [ ]
                                {
                                    intI ,                                      // Format Item 0: Detail record {0}
                                    strEmbeddResourceFileName ,                 // Format Item 1: internal resource {1}
                                    EXPECTED_FIELD_COUNT ,                      // Format Item 2: Expected field count = {2}
                                    astrFields.Length ,                         // Format Item 3: Actual field count   = {3}
                                    astrAllMapItems [ intI ] ,                  // Format Item 4: Actual record        = {4}
                                    Environment.NewLine                         // Format Item 5: Platform-specific newline
                                } ) );
                    }   // FALSE (unanticipated outcome) block, if ( astrFields.Length == EXPECTED_FIELD_COUNT )
                }   // FALSE (detail row) block, if ( intI == ArrayInfo.ARRAY_FIRST_ELEMENT )
            }   // for ( int intI = ArrayInfo.ARRAY_FIRST_ELEMENT ; intI < astrAllMapItems.Length ; intI++ )

            return rFunctionMaps;
        }   // private StringFixups.StringFixup [ ] GetSStringFixups


        private FunctionMapItem [ ] _map = null;

  
        private RestClient _restClient = null;


        private StringFixups _responseStringFixups = null;


        private SymbolInfo [ ] _symbolInfo = null;


        private class FunctionMapItem : IComparable<FunctionMapItem>
        {
            public ApiFunction function
            {
                get; set;
            }
            public string parameter
            {
                get; set;
            }

            int IComparable<FunctionMapItem>.CompareTo ( FunctionMapItem that )
            {
                return function.CompareTo ( that.function );
            }   // int IComparable<FunctionMapItem>.CompareTo


            public override bool Equals ( object obj )
            {
                FunctionMapItem that = obj as FunctionMapItem;
                return function.Equals ( that.function );
            }   // public override bool Equals


            public override int GetHashCode ( )
            {
                return function.GetHashCode ( );
            }   // public override int GetHashCode


            public override string ToString ( )
            {
                return string.Format (
                    Properties.Resources.TPL_API_FUNCTION_MAP_TOSTRING ,        // Format control string
                    function ,                                                  // Format Item 0: Function = {0}, Parameter = {1}
                    parameter );                                                // Format Item 1: Parameter = {1}
            }   // public override string ToString
        }   // private class FunctionMap
    }   // class StockTickerEngine is implicitly internal.
}   // partial namespace StockTickerSparklines