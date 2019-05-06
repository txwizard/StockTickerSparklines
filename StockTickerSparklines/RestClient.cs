using System;

using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

using WizardWrx;

namespace StockTickerSparklines
{
    /// <summary>
    /// This class is a wrapper for the 
    /// </summary>
    public class RestClient
    {
        #region Enumerations
        /// <summary>
        /// Reducing the four verbs to an enumeration reduces a string pointer
        /// and its memory to a simple value type.
        /// </summary>
        public enum HttpVerb
        {
            /// <summary>
            /// The simplest verb supports Read (Inquire) tasks, covering
            /// the "R" element of the CRUD model. This verb frequently gets its
            /// data from a query string.
            /// </summary>
            Get,

            /// <summary>
            /// Post is the usual mechanism for creating new records, covering
            /// the "C" element of the CRUD model. This verb frequently requires
            /// a payload of some sort, since its inputs are usually more than
            /// is prudent to incorporate into a query string, not to mention
            /// that it is usually too sensitive for that to be wise.
            /// </summary>
            Post,

            /// <summary>
            /// Put is the usual mechanism for updating existing records,
            /// covering the "U" element of the CRUD model. Like POST, PUT is
            /// usually accompanied by a payload.
            /// </summary>
            Put,

            /// <summary>
            /// Delete is the usual mechanism for deleting records from a table,
            /// covering the "D" element of the CRUD model. The delete verb may
            /// be implemented by way of a query string, although it may be best
            /// to put something magic into the payload that must be present to
            /// make drive-by attacks harder to pull off.
            /// </summary>
            Delete
        }   // public enum HttpVerb


        /// <summary>
        /// Both responses and payloads usually require a MIME type to be
        /// specified. Specifying the mime type as an enumeration provides the
        /// same benefits as are afforded by the HttpVerb enumeration.
        /// </summary>
        /// <see href="https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Complete_list_of_MIME_types"/>
        public enum ContentType
        {
            /// <summary>
            /// Common Name:    Comma-separated values (CSV), 
            /// MIME Type:      text/csv
            /// File Extension: .csv
            /// </summary>
            CSV,

            /// <summary>
            /// Common Name:    JSON (JavaScript Object Notation)
            /// MIME Type:      application/json
            /// File Extension: .json
            /// </summary>
            JSON,
            /// Common Name:    eXtensible Markup Language (XML)
            /// MIME Type:      application/xml if not readable by casual users (RFC 3023, section 3)
            ///                 text/xml if readable by casual users (RFC 3023, section 3)
            /// File Extension: .xml
            XML,
        }   // public enum MimeType
        #endregion // Enumerations


        #region Properties
        /// <summary>
        /// Gets or sets the EndPoint (URI) through which the API is accessed
        /// </summary>
        public string EndPoint
        {
            get; set;
        }   // EndPoint property (Read/Write)


        /// <summary>
        /// Gets or sets the HTTP verb in terms of the HttpVerb enumeration,
        /// which is freely convertible to its commonly used string form
        /// </summary>
        public HttpVerb Verb
        {
            get; set;
        }   // Verb property (Read/Write)


        /// <summary>
        /// Gets the string representation of the HttpVerb property
        /// </summary>
        public string VerbString
        {
            get
            {
                return s_astrVerbStrings [ ( int ) Verb ];
            }   // public string VerbString property getter method
        }   // public string VerbString property


        /// <summary>
        /// Gets the Mime type in terms of the ContentType enumeration,
        /// which is freely convertible to its commonly used string form
        /// </summary>
        public ContentType MimeType
        {
            get; set;
        }   // MimeType property (Read/Write)


        /// <summary>
        /// Gets the MIME type string that corresponds to the MimeType property
        /// </summary>
        public string MimeTypeString
        {
            get
            {
                return s_astrMimeTypes [ ( int ) MimeType ];
            }   // MimeTypeString property getter method
        }   // MimeTypeString property


        /// <summary>
        /// Gets or sets the payload of the request
        /// </summary>
        public string PostData
        {
            get; set;
        }   // PostData property
        #endregion  // Properties


        #region Constructors
        public RestClient ( )
        {
            InitializeInstance (
                SpecialStrings.EMPTY_STRING ,               // string       pstrEndPoint
                HttpVerb.Get ,                              // HttpVerb     penmVerb
                ContentType.JSON ,                          // ContentType  penmType
                SpecialStrings.EMPTY_STRING );              // string       pstrPostData
        }   // RestClient constructor 1 of 4


        public RestClient (
            string pstrEndPoint )
        {
            InitializeInstance (
                pstrEndPoint ,                              // string       pstrEndPoint
                HttpVerb.Get ,                              // HttpVerb     penmVerb
                ContentType.JSON ,                          // ContentType  penmType
                SpecialStrings.EMPTY_STRING );              // string       pstrPostData
        }   // // RestClient constructor 2 of 4


        public RestClient (
            string pstrEndPoint ,
            HttpVerb penmVerb )
        {
            InitializeInstance (
                pstrEndPoint ,                              // string       pstrEndPoint
                penmVerb ,                                  // HttpVerb     penmVerb
                ContentType.JSON ,                          // ContentType  penmType
                SpecialStrings.EMPTY_STRING );              // string       pstrPostData
        }   // RestClient constructor 3 of 4


        public RestClient (
            string pstrEndPoint ,
            HttpVerb penmVerb ,
            ContentType penmContentType )
        {
            InitializeInstance (
                pstrEndPoint ,                              // string       pstrEndPoint
                penmVerb ,                                  // HttpVerb     penmVerb
                penmContentType ,                           // ContentType  penmType
                SpecialStrings.EMPTY_STRING );              // string       pstrPostData
        }   // RestClient constructor 4 of 5


        public RestClient (
            string pstrEndPoint ,
            HttpVerb penmVerb ,
            string pstrPostData )
        {
            InitializeInstance (
                pstrEndPoint ,                              // string       pstrEndPoint
                penmVerb ,                                  // HttpVerb     penmVerb
                ContentType.JSON ,                          // ContentType  penmType
                pstrPostData );                             // string       pstrPostData
        }   // RestClient constructor 5 of 5


        private void InitializeInstance (
            string pstrEndPoint ,
            HttpVerb penmVerb ,
            ContentType penmType ,
            string pstrPostData )
        {
            EndPoint = pstrEndPoint;
            Verb = penmVerb;
            MimeType = penmType;
            PostData = pstrPostData;
        }   // InitializeInstance
        #endregion // Constructors


        #region MakeRequest Methods
        /// <summary>
        /// Make a request based entirely on the instance properties.
        /// </summary>
        /// <returns>
        /// If the method succeeds, the return value is the response returned by
        /// the endpoint. Otherwise, the return value is the empty string.
        /// </returns>
        public string MakeRequest ( )
        {
            List<string> lstHeaders = new List<string> ( );
            return MakeRequest (
                WizardWrx.SpecialStrings.EMPTY_STRING ,     // QueryString, formatted as a query string (none)
                lstHeaders ,                                // Headers, as a list of strings (empty)
                Verb );                                     // HTTP Verb (per instance property)
        }   // public string MakeRequest method (1 of 5)


        /// <summary>
        /// Make a request based on the instance properties and the QueryString
        /// specified by argument <paramref name="pstrQueryString"/>.
        /// </summary>
        /// <param name="pstrQueryString">
        /// QueryString must be a well-formed queryString composed of key-value
        /// pairs, each composed of a name and a value, separated by an equals
        /// sign. The first name must be preceded by a question mark, with the
        /// subsequent names preceded by an ampersand.
        /// </param>
        /// <returns>
        /// If the method succeeds, the return value is the response returned by
        /// the endpoint. Otherwise, the return value is the empty string.
        /// </returns>
        public string MakeRequest ( string pstrQueryString )
        {
            List<string> lstHeaders = new List<string> ( );
            return MakeRequest (
                pstrQueryString ,                           // QueryString, formatted as a query string (none)
                lstHeaders ,                                // Headers, as a list of strings (empty)
                Verb );                                     // HTTP Verb (per instance property)
        }   // public string MakeRequest method (2 of 5)


        /// <summary>
        /// Make a request based on the properties of the instance and the
        /// specified list of <paramref name="plstHeaders"/>.
        /// </summary>
        /// <param name="plstHeaders">
        /// The headers collection is submitted as a generic List of strings,
        /// each of which is a name-value pair delimited by an equals sign.
        /// </param>
        /// <returns>
        /// If the method succeeds, the return value is the response returned by
        /// the endpoint. Otherwise, the return value is the empty string.
        /// </returns>
        public string MakeRequest ( List<string> plstHeaders )
        {
            return MakeRequest (
                SpecialStrings.EMPTY_STRING ,               // QueryString, formatted as a query string (none)
                plstHeaders ,                               // Headers, as a list of strings (per argument)
                Verb );                                     // HTTP Verb (per instance property)
        }   // public string MakeRequest method (3 of 5)


        /// <summary>
        /// Make a request based on the properties of the instance, modified by
        /// the specified <paramref name="pstrQueryString"/> and <paramref name="plstHeaders"/>.
        /// </summary>
        /// <param name="pstrQueryString">
        /// QueryString must be a well-formed queryString composed of key-value
        /// pairs, each composed of a name and a value, separated by an equals
        /// sign. The first name must be preceded by a question mark, with the
        /// subsequent names preceded by an ampersand.
        /// </param>
        /// <param name="plstHeaders">
        /// The headers collection is submitted as a generic List of strings,
        /// each of which is a name-value pair delimited by an equals sign.
        /// </param>
        /// <returns>
        /// If the method succeeds, the return value is the response returned by
        /// the endpoint. Otherwise, the return value is the empty string.
        /// </returns>
        public string MakeRequest (
            string pstrQueryString ,
            List<string> plstHeaders )
        {
            return MakeRequest (
                pstrQueryString ,                           // QueryString, formatted as a query string (per agrument list)
                plstHeaders ,                               // Headers, as a list of strings (per argument list)
                Verb );                                     // HTTP Verb (per instance property)
        }   // public string MakeRequest method (4 of 5)


        /// <summary>
        /// Make a request based on the properties of the instance, modified by
        /// the specified <paramref name="pstrQueryString"/>, <paramref name="plstHeaders"/>,
        /// and <paramref name="httpVerb"/>.
        /// </summary>
        /// <param name="pstrQueryString">
        /// QueryString must be a well-formed queryString composed of key-value
        /// pairs, each composed of a name and a value, separated by an equals
        /// sign. The first name must be preceded by a question mark, with the
        /// subsequent names preceded by an ampersand.
        /// </param>
        /// <param name="plstHeaders">
        /// The headers collection is submitted as a generic List of strings,
        /// each of which is a name-value pair delimited by an equals sign.
        /// </param>
        /// <param name="httpVerb">
        /// The verb is specified in terms of the HttpVerb enumeration, and is
        /// mapped internally to the appropriate string.
        /// </param>
        /// <returns>
        /// If the method succeeds, the return value is the response returned by
        /// the endpoint. Otherwise, the return value is the empty string.
        /// </returns>
        public string MakeRequest (
            string pstrQueryString ,                        // QueryString, formatted as a query string
            List<string> headers ,                          // Headers, as a list of strings
            HttpVerb httpVerb )                             // HTTP Verb, per locally defined enumeration
        {
            var request = ( HttpWebRequest ) WebRequest.Create ( EndPoint + pstrQueryString );

            request.Method = s_astrVerbStrings [ ( int ) Verb ];
            request.ContentLength = MagicNumbers.EMPTY_STRING_LENGTH;
            request.ContentType = s_astrMimeTypes [ ( int ) MimeType ];

            request.Accept = "*/*";                         // Accept anything.

            foreach ( string header in headers )
            {
                request.Headers.Add ( header );
            }   // foreach ( string header in headers )

            if ( !string.IsNullOrEmpty ( PostData ) && Verb == HttpVerb.Post )
            {
                Encoding encoding = new UTF8Encoding ( );
                byte [ ] abytBytes = Encoding.GetEncoding (
                    "iso-8859-1" ).GetBytes (
                    PostData );
                request.ContentLength = abytBytes.Length;

                using ( Stream writeStream = request.GetRequestStream ( ) )
                {
                    writeStream.Write (
                        abytBytes ,
                        ListInfo.BEGINNING_OF_BUFFER ,
                        abytBytes.Length );
                }   // using ( Stream writeStream = request.GetRequestStream ( ) )
            }   // if ( !string.IsNullOrEmpty ( PostData ) && Method == HttpVerb.POST )

            using ( HttpWebResponse response = ( HttpWebResponse ) request.GetResponse ( ) )
            {
                string responseValue = string.Empty;

                if ( response.StatusCode != HttpStatusCode.OK )
                {
                    string strMessage = string.Format (
                        "Request failed. Received HTTP response code {0}" ,
                        response.StatusCode );
                    throw new ApplicationException ( strMessage );
                }   // if ( response.StatusCode != HttpStatusCode.OK )

                //  ------------------------------------------------------------
                //  Grab the response.
                //  ------------------------------------------------------------

                using ( var responseStream = response.GetResponseStream ( ) )
                {
                    if ( responseStream != null )
                    {
                        using ( StreamReader reader = new StreamReader ( responseStream ) )
                        {
                            responseValue = reader.ReadToEnd ( );
                        }   // using ( StreamReader reader = new StreamReader ( responseStream ) )
                    }   // if ( responseStream != null )
                }   // using ( var responseStream = response.GetResponseStream ( ) )

                return responseValue;
            }   // using ( var response = ( HttpWebResponse ) request.GetResponse ( ) )
        }   // // public string MakeRequest method (5 of 5)
        #endregion // MakeRequest methods


        #region Static read-only lookup arrays for mapping enumerations to strings
        static readonly string [ ] s_astrVerbStrings =
        {
            @"GET" ,                                        // HttpVerb.Get
            @"POST" ,                                       // HttpVerb.Post
            @"PUT" ,                                        // HttpVerb.Put
            @"DELETE"                                       // HttpVerb.Delete
        };  // s_astrVerbStrings

        static readonly string [ ] s_astrMimeTypes =
        {
            @"text/csv" ,                                   // ContentType.CSV
            @"text/json" ,                                  // ContentType.JSON
            @"application/xml"                              // ContentType.XML
        };  // s_astrMimeTypes


        static readonly string [ ] s_astrMimeDocumentExtensions =
        {
            @".csv" ,                                       // ContentType.CSV
            @".json" ,                                      // ContentType.JSON
            @".xml"                                         // ContentType.XML
        };  // s_astrMimeDocumentExtensions
        #endregion  // Static read-only lookup arrays for mapping enumerations to strings


        public class ErrorResponse
        {
            /// <summary>
            /// Static method ResponseIsErrorMessage uses this constructed string to
            /// identify a response as an error message.
            /// </summary>
            public const string RESPONSE_IS_ERROR = SpecialStrings.DOUBLE_QUOTE + @"Error Message" + SpecialStrings.DOUBLE_QUOTE + SpecialStrings.COLON;


            public string ErrorMessage
            {
                get; set;
            }   // public string ErrorMessage property (Read/Write)


            /// <summary>
            /// Evaluate the raw response string, returning TRUE if it indicates
            /// that the response is an error message.
            /// </summary>
            /// <param name="pstrResponse">
            /// String containing the response returned by the API
            /// </param>
            /// <returns>
            /// TRUE if the response is an error, otherwise FALSE
            /// </returns>
            public static bool ResponseIsErrorMessage ( string pstrResponse )
            {
                return pstrResponse.IndexOf ( RESPONSE_IS_ERROR ) > ListInfo.INDEXOF_NOT_FOUND;
            }   // public static bool ResponseIsErrorMessage
        }   // public class ErrorResponse


        public class QueryStringBuilder
        {
            public QueryStringBuilder ( )
            {
            }   // public QueryStringBuilder constructor


            public void AddParameter (
                string pstrParamName ,
                string pstrParamValue )
            {
                if ( string.IsNullOrEmpty ( pstrParamName ) )
                    throw new ArgumentNullException ( nameof ( pstrParamName ) );

                if ( string.IsNullOrEmpty ( pstrParamValue ) )
                    throw new ArgumentNullException ( nameof ( pstrParamValue ) );

                StringBuilder paramBuilder = new StringBuilder ( MagicNumbers.CAPACITY_MAX_PATH );

                paramBuilder.Append ( _queryStringBuilder == null ? SpecialCharacters.QUESTION_MARK : SpecialCharacters.AMPERSAND );
                paramBuilder.Append ( pstrParamName );
                paramBuilder.Append ( SpecialCharacters.EQUALS_SIGN );
                paramBuilder.Append ( pstrParamValue );

                _queryStringBuilder = _queryStringBuilder ?? new StringBuilder ( MagicNumbers.CAPACITY_01KB );
                _queryStringBuilder.Append ( paramBuilder );
            }   // public void AddParameter


            public string GetQueryString ( )
            {
                return _queryStringBuilder.ToString ( );
            }   // public string GetQueryString
            private StringBuilder _queryStringBuilder = null;
        }   // public class QueryStringBuilder
    }   // public class RestClient
}   // partial namespace StockTickerSparklines