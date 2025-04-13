*!*	Tokenizer - a VFP class to extract tokens from a source string

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

#DEFINE	SAFETHIS		ASSERT !USED("This") AND TYPE("This") == "O"

DEFINE CLASS Tokenizer AS Custom

	* a collection of patterns to identify tokens
	ADD OBJECT TokenPatterns AS Collection
	* a collection of identified tokens
	ADD OBJECT Tokens AS Collection

	* a regular expression engine
	HIDDEN RegExpr
	RegExpr = .NULL.
	* and its class
	RegExprClass = "VBScript.RegExp"

	* where an error occurred in the string (no patterns available)
	ErrorPointer = ""
	* where the tokenizer stopped parsing (no patterns available due to filtering)
	StopPointer = ""
	* force the pattern to start at the source beginning
	ForceStart = .T.
	* trim the source before parsing it (hopefully, this will simplify the creation of patterns)
	TrimBeforeParse = .T.
	* ignore case while matching the expressions
	IgnoreCase = .F.

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="errorpointer" type="property" display="ErrorPointer"/>' + ;
						'<memberdata name="forcestart" type="property" display="ForceStart"/>' + ;
						'<memberdata name="ignorecase" type="property" display="IgnoreCase"/>' + ;
						'<memberdata name="regexprclass" type="property" display="RegExprClass"/>' + ;
						'<memberdata name="stoppointer" type="property" display="StopPointer"/>' + ;
						'<memberdata name="tokenpatterns" type="property" display="TokenPatterns"/>' + ;
						'<memberdata name="tokens" type="property" display="Tokens"/>' + ;
						'<memberdata name="trimbeforeparse" type="property" display="TrimBeforeParse"/>' + ;
						'<memberdata name="addtokenpattern" type="method" display="AddTokenPattern"/>' + ;
						'<memberdata name="gettokens" type="method" display="GetTokens"/>' + ;
						"</VFPData>"

	PROCEDURE Init (RXEngine AS String)
		IF PCOUNT() == 1
			This.RegExprClass = m.RXEngine
		ELSE
			This.SetRegExpr()
		ENDIF
	ENDPROC

	PROCEDURE Destroy
		This.RegExpr = .NULL.
	ENDPROC

	HIDDEN PROCEDURE SetRegExpr ()
		This.RegExpr = CREATEOBJECT(This.RegExprClass)
		This.RegExpr.Ignorecase = This.IgnoreCase
	ENDPROC

	HIDDEN PROCEDURE RegExprClass_Assign (NewValue AS String)
		IF VARTYPE(m.NewValue) == "C"
			This.RegExprClass = m.NewValue
			This.SetRegExpr()
		ENDIF
	ENDPROC

	HIDDEN PROCEDURE IgnoreCase_Assign (NewValue AS Logical)
		IF VARTYPE(m.NewValue) == "L"
			This.IgnoreCase = m.NewValue
			This.RegExpr.IgnoreCase = m.NewValue
		ENDIF
	ENDPROC
	
	* AddTokenPattern
	* add a pattern to the pattern collection
	* patterns are verified by the order they were added to the collection
	* a pattern should include at least 2 capturing groups, and must include at least one: one to fetch the value of the token;
	*  the other to store the actual source that was read and may be skipped for the next step of the parsing process
	FUNCTION AddTokenPattern (Pattern AS String, Type AS String, ValueGroup AS Integer, SkipGroup AS Integer)
	
		LOCAL NewTokenType AS TokenType
		LOCAL NextPattern AS String

		* make sure the pattern points to the beginning of the source string
		m.NextPattern = IIF(This.ForceStart AND LEFT(m.Pattern, 1) != "^", "^", "") + m.Pattern

		* create a TokenType object
		IF PCOUNT() > 2
			m.NewTokenType = CREATEOBJECT("TokenType", m.NextPattern, m.Type, m.ValueGroup, m.SkipGroup)
		ELSE
			m.NewTokenType = CREATEOBJECT("TokenType", m.NextPattern, m.Type)
		ENDIF

		* and it to the collection
		This.TokenPatterns.Add(m.NewTokenType)

	ENDFUNC

	* GetTokens
	* get as much tokens as it can
	FUNCTION GetTokens (SourceString AS String, AllowedTypes AS String) AS Logical

		SAFETHIS

		LOCAL Source AS String
		LOCAL ARRAY Allowed(1)
		LOCAL Filtered AS Logical

		LOCAL TokenType AS TokenType
		LOCAL Token AS Token
		LOCAL TokenIndex AS Integer
		LOCAL TokenText AS String

		LOCAL Matches AS VBScript.MatchCollection

		* remove surrounding white space 
		IF This.TrimBeforeParse
			m.Source = ALLTRIM(m.SourceString, 0, " ", CHR(13), CHR(10), CHR(9))
		ELSE
			m.Source = m.SourceString
		ENDIF

		* if a comma-separated list of token types were passed, filtering will be on
		* otherwise, the source must be completely parsed
		IF PCOUNT() = 2
			ALINES(m.Allowed, m.AllowedTypes, 1 + 4, ",")
			m.Filtered = .T.
		ELSE
			m.Filtered = .F.
		ENDIF

		* clean the identified tokens collection
		This.Tokens.Remove(-1)
		* prepare the parsing
		This.ErrorPointer = ""
		This.StopPointer = ""

		* while there is something in the source to parse
		DO WHILE !EMPTY(m.Source)

			* go through all defined / filtered tokens
			m.TokenText = .NULL.
			m.TokenIndex = 1

			DO WHILE ISNULL(m.TokenText) AND m.TokenIndex <= This.TokenPatterns.Count

				m.TokenType = This.TokenPatterns.Item(m.TokenIndex)
				* if all types are to be verified, or if the current one is one of the allowed types
				IF !m.Filtered OR ASCAN(m.Allowed, m.TokenType.Type, -1, -1, -1, 1 + 2 + 4) != 0

					* verify the expression against (the beginning of) the source
					This.RegExpr.Pattern = m.TokenType.Pattern
					m.Matches = This.RegExpr.Execute(m.Source)

					* a match was found, and we matching the whole pattern or there are enough groups to identify the token value
					IF m.Matches.Count = 1 ;
							AND (m.TokenType.ValueGroup < 0 OR ;
								(m.Matches.Item(0).SubMatches.Count > MAX(m.TokenType.ValueGroup, m.TokenType.SkipGroup) ;
								AND !ISNULL(m.Matches.Item(0).SubMatches(m.TokenType.ValueGroup)) ;
								AND !ISNULL(m.Matches.Item(0).SubMatches(m.TokenType.SkipGroup))))

						* matching the whole pattern?
						IF m.TokenType.ValueGroup < 0
							* get the token
							m.TokenText = m.Matches.Item(0).Value
							* and skip to the next point in the source
							m.Source = SUBSTR(m.Source, m.Matches.Item(0).FirstIndex + m.Matches.Item(0).Length + 1)
						ELSE
							* get the token
							m.TokenText = m.Matches.Item(0).SubMatches(m.TokenType.ValueGroup)
							* and skip the source to the next token
							m.Source = SUBSTR(m.Source, LEN(m.Matches.Item(0).SubMatches(m.TokenType.SkipGroup)) + 1)
						ENDIF
						* mark the stop, in case there are no more available patterns for the rest of the string
						This.StopPointer = m.Source

						* trim the white space, if needed
						IF This.TrimBeforeParse
							m.Source = LTRIM(m.Source, 0, " ", CHR(13), CHR(10), CHR(9))
						ENDIF

						* a token was found
						m.Token = CREATEOBJECT("Token", m.TokenText, m.TokenType.Type)
						* add it to the collection of identified tokens
						This.Tokens.Add(m.Token)

					ENDIF

				ENDIF

				* go to the next token (in case a match wasn't found)
				m.TokenIndex = m.TokenIndex + 1

			ENDDO

			* in the last run we didn't find any token?
			IF ISNULL(m.TokenText)

				* return no-success, unless filtering was on and at least one token was found previously
				IF !m.Filtered OR This.Tokens.Count = 0
					This.ErrorPointer = m.Source
					RETURN .F.
				ELSE
					RETURN .T.
				ENDIF

			ENDIF

		ENDDO

		* success: the tokens may be fetched from the tokens collection
		RETURN .T.
	ENDFUNC

ENDDEFINE

* A type of token

* In a collection of token types, different patterns may share the same Type (for alternative token syntaxes).
* The whoie pattern may be tested (when .ValueGroup = -1) or may be singled out in a group (when .ValueGroup >= 0)
* In this case, at least one group must be defined, to discriminate the group that will match the token value,
*  and other to match the part of the source string that may be skipped.
DEFINE CLASS TokenType AS Custom

	Pattern = ""
	Type = ""
	SkipGroup = 0
	ValueGroup = -1

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="pattern" type="property" display="Pattern"/>' + ;
						'<memberdata name="skipgroup" type="property" display="SkipGroup"/>' + ;
						'<memberdata name="type" type="property" display="Type"/>' + ;
						'<memberdata name="valuegroup" type="property" display="ValueGroup"/>' + ;
						"</VFPData>"

	FUNCTION Init (Pattern AS String, Type AS String, ValueGroup AS Integer, SkipGroup AS Integer)
	
		This.Pattern = m.Pattern
		This.Type = m.Type
		IF PCOUNT() > 2
			This.ValueGroup = m.ValueGroup
			This.SkipGroup = m.SkipGroup
		ENDIF

	ENDFUNC

ENDDEFINE

* a token that was identified during the parsing (and to what type it belongs)

DEFINE CLASS Token AS Custom

	Value = ""
	Type = ""

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="type" type="property" display="Type"/>' + ;
						'<memberdata name="value" type="property" display="Value"/>' + ;
						"</VFPData>"

	FUNCTION Init (Value AS String, Type AS String)
	
		This.Value = m.Value
		This.Type = m.Type

	ENDFUNC

ENDDEFINE
