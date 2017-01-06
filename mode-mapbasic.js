ace.define("ace/mode/vbscript_highlight_rules",["require","exports","module","ace/lib/oop","ace/mode/text_highlight_rules"], function(require, exports, module) {
"use strict";

var oop = require("../lib/oop");
var TextHighlightRules = require("./text_highlight_rules").TextHighlightRules;

var MapBasicHighlightRules = function() {

    var keywordMapper = this.createKeywordMapper({
        "keyword.control.asp":  "For|To|Next|Exit|Do|Loop|While|Wend|If|Then|ElseIf|Case|GoTo|End"
			+ "End Program|Terminate Application|End MapInfo|OnError|Resume|Function|Sub",
            
			//VBS - keword.control.asp
			//"If|Then|Else|ElseIf|End|While|Wend|For|To|Each|Case|Select|Return"
			//+ "|Continue|Do|Until|Loop|Next|With|Exit|Function|Property|Type|Enum|Sub|IIf",
			
        "storage.type.asp": "Dim|Call|Define|Declare|Redim|Randomize|Type",
		
			//VBS - storage.type.asp
			//"Dim|Call|Class|Const|Dim|Redim|Set|Let|Get|New|Randomize|Option|Explicit"
		
        "storage.modifier.asp": "",
		
			//VBS - storage.modifier.asp
			//"Private|Public|Default"
		
        "keyword.operator.asp": "Mod|And|Not|Or|as|Contains|Part|Entire|Within|Partly|Entirely|Intersects",
		
			//VBS - keyword.operator.asp
			//"Mod|And|Not|Or|Xor|as"
		
		//CONSTANTS
		// There are many constants in MapBasic and the list goes on and on but all require MapBasic.DEF. Do we include? - Probably
		// Might be best to create on the fly from MapBasic.DEF etc. --- REGEX: "Define\s+\b(\w*)\b" return "$1"
        "constant.language.asp": "$REPLACE_CONSTANTS$",
		
			//VBS - constant.language.asp
			//"Empty|False|Nothing|Null|True"
		
        "support.class.asp": "",
			//VBS - support.class.asp
			//"Application|ObjectContext|Request|Response|Server|Session"
			
        "support.class.collection.asp": "Contents|StaticObjects|ClientCertificate|Cookies|Form|QueryString|ServerVariables",
        
		"support.constant.asp": "TotalBytes|Buffer|CacheControl|Charset|ContentType|Expires|ExpiresAbsolute"
            + "|IsClientConnected|PICS|Status|ScriptTimeout|CodePage|LCID|SessionID|Timeout",
        
		"support.function.asp": "Handler"
				
				//VBS
				//"Lock|Unlock|SetAbort|SetComplete|BinaryRead|AddHeader|AppendToLog"
				//+ "|BinaryWrite|Clear|Flush|Redirect|Write|CreateObject|HTMLEncode|MapPath|URLEncode|Abandon|Convert|Regex",
        
		"support.function.event.asp": "SelChangedHandler|WinClosedHandler|WinChangedHandler|WinFocusChangedHandler"
			+ "|RemoteMsgHandler|RemoteQueryHandler|RemoteMapGenHandler|ToolHandler|EndHandler|ForegroundTaskSwitchHandler"
			//VBS
			//"Application_OnEnd|Application_OnStart"
            //+ "|OnTransactionAbort|OnTransactionCommit|Session_OnEnd|Session_OnStart",
			
        
		//List of functions/procs/etc from MapBasic Reference using RegEx: "(.*)(function|procedure|statement|clause)"
		"support.function.vb.asp": "Array|Add|Asc|Atn|CBool|CByte|CCur|CDate|CDbl|Chr|CInt|CLng"
            + "|Conversions|Cos|CreateObject|CSng|CStr|Date|DateAdd|DateDiff|DatePart|DateSerial"
            + "|DateValue|Day|Derived|Math|Escape|Eval|Exists|Exp|Filter|FormatCurrency"
            + "|FormatDateTime|FormatNumber|FormatPercent|GetLocale|GetObject|GetRef|Hex"
            + "|Hour|InputBox|InStr|InStrRev|Int|Fix|IsArray|IsDate|IsEmpty|IsNull|IsNumeric"
            + "|IsObject|Item|Items|Join|Keys|LBound|LCase|Left|Len|LoadPicture|Log|LTrim|RTrim"
            + "|Trim|Maths|Mid|Minute|Month|MonthName|MsgBox|Now|Oct|Remove|RemoveAll|Replace"
            + "|RGB|Right|Rnd|Round|ScriptEngine|ScriptEngineBuildVersion|ScriptEngineMajorVersion"
            + "|ScriptEngineMinorVersion|Second|SetLocale|Sgn|Sin|Space|Split|Sqr|StrComp|String|StrReverse"
            + "|Tan|Time|Timer|TimeSerial|TimeValue|TypeName|UBound|UCase|Unescape|VarType|Weekday|WeekdayName|Year",
        
		"support.type.vb.asp": "SmallInt|Integer|Float|String|Logical|Date|Object|Alias|Pen|Brush"
			
			//VBS
			// "support.type.vb.asp": "vbtrue|vbfalse|vbcr|vbcrlf|vbformfeed|vblf|vbnewline|vbnullchar|vbnullstring|"
            // + "int32|vbtab|vbverticaltab|vbbinarycompare|vbtextcomparevbsunday|vbmonday|vbtuesday|vbwednesday"
            // + "|vbthursday|vbfriday|vbsaturday|vbusesystemdayofweek|vbfirstjan1|vbfirstfourdays|vbfirstfullweek"
            // + "|vbgeneraldate|vblongdate|vbshortdate|vblongtime|vbshorttime|vbobjecterror|vbEmpty|vbNull|vbInteger"
            // + "|vbLong|vbSingle|vbDouble|vbCurrency|vbDate|vbString|vbObject|vbError|vbBoolean|vbVariant"
            // + "|vbDataObject|vbDecimal|vbByte|vbArray"
			
    }, "identifier", true);

    this.$rules = {
    "start": [
        {
            token: [
                "meta.ending-space"
            ],
            regex: "$"
        },
        {
            token: [null],
            regex: "^(?=\\t)",
            next: "state_3"
        },
        {
            token: [null],
            regex: "^(?= )",
            next: "state_4"
        },
        {
            token: [
                "text",
                "storage.type.function.asp",
                "text",
                "entity.name.function.asp",
                "text",
                "punctuation.definition.parameters.asp",
                "variable.parameter.function.asp",
                "punctuation.definition.parameters.asp"
            ],
			regex: "^(\\s*)(Function|Sub)(\\s+)([a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*)(\\s*)(\\()([^)]*)(\\))"
            //MB (ASSUMED!)
			// ^(\\s*)(Function|Sub)(\\s+)([a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*)(\\s*)(\\()([^)]*)(\\))
			//VBS
			// regex: "^(\\s*)(Function|Sub)(\\s+)([a-zA-Z_]\\w*)(\\s*)(\\()([^)]*)(\\))"
        },
        {
            token: "punctuation.definition.comment.asp",
            regex: "'",
				//MB
				//'
				//VBS
				//"'|REM(?=\\s|$)",
				
			next: "comment",
            caseInsensitive: true
        },
        {
            token: "storage.type.asp",
            regex: "OnError GoTo",
				//MB
				//OnError GoTo {label | 0}
				//VBS
				//regex: "On Error Resume Next|On Error GoTo",
				
            caseInsensitive: true
        },
        {
            token: "punctuation.definition.string.begin.asp",
            regex: '"',
            next: "string"
        },
        {
            token: [
                "punctuation.definition.variable.asp"
            ],
            regex: "[a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*"	//Innacurately allows a@a as a variable name
				//MB
				//[a-zA-Z_~\x80-\xFF][0-9@&%$#!\x0C\x09a-zA-Z_~\x80-\xFF]*
				//VBS
				//regex: "(\\$)[a-zA-Z_x7f-xff][a-zA-Z0-9_x7f-xff]*?\\b\\s*"
        },
        {
            token: "constant.numeric.asp",
            regex: "-?(?:(?:\\&[Hh][0-9a-fA-F]*)|\\b(?:(?:[0-9]+\\.?[0-9]*)|(?:\\.[0-9]+))(?:(?:e|E)(?:\\+|-)?[0-9]+)?)\\b"
				//MB (Examples: "1","1.1",".1","&h1A", "1E10", "1E-10")
				//-?(?:(?:\&[Hh][0-9a-fA-F]*)|\b(?:(?:[0-9]+\.?[0-9]*)|(?:\.[0-9]+))(?:(?:e|E)(?:\+|-)?[0-9]+)?)\b
				//VBS
				//regex: "-?\\b(?:(?:0(?:x|X)[0-9a-fA-F]*)|(?:(?:[0-9]+\\.?[0-9]*)|(?:\\.[0-9]+))(?:(?:e|E)(?:\\+|-)?[0-9]+)?)(?:L|l|UL|ul|u|U|F|f)?\\b"
        },
        {
            regex: "\\w+",
            token: keywordMapper
        },
        {
            token: ["entity.name.function.asp"],
            regex: "(?:([a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*)(?=\\(\\)?))"	//Assumed
				//MB
				//[a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*\(\)
				//VBS
				//regex: "(?:(\\b[a-zA-Z_x7f-xff][a-zA-Z0-9_x7f-xff]*?\\b)(?=\\(\\)?))"
        },
        {
            token: ["keyword.operator.asp"],
            regex: "\\-|\\+|\\*|\\/|\\>|\\<|\\=|\\&(?!h[0-9a-fA-F]*\\b)"
				//MB (surpress error where hex value)
				// \-|\+|\*|\/|\>|\<|\=|\&(?!h[0-9a-fA-F]*\b)
				//VBS
				//regex: "\\-|\\+|\\*\\/|\\>|\\<|\\=|\\&"
			
        }
    ],
    "state_3": [
        {
            token: [
                "meta.odd-tab.tabs",
                "meta.even-tab.tabs"
            ],
            regex: "(\\t)(\\t)?"
        },
        {
            token: "meta.leading-space",
            regex: "(?=[^\\t])",
            next: "start"
        },
        {
            token: "meta.leading-space",
            regex: ".",
            next: "state_3"
        }
    ],
    "state_4": [
        {
            token: ["meta.odd-tab.spaces", "meta.even-tab.spaces"],
            regex: "(  )(  )?"
        },
        {
            token: "meta.leading-space",
            regex: "(?=[^ ])",
            next: "start"
        },
        {
            defaultToken: "meta.leading-space"
        }
    ],
    "comment": [
        {
            token: "comment.line.apostrophe.asp",
            regex: "$",
			
			//VBS
			//regex: "$|(?=(?:%>))",
			
            next: "start"
        },
        {
            defaultToken: "comment.line.apostrophe.asp"
        }
    ],
    "string": [
        {
            token: "constant.character.escape.apostrophe.asp",
            regex: '""'
        },
        {
            token: "string.quoted.double.asp",
            regex: '"',
            next: "start"
        },
        {
            defaultToken: "string.quoted.double.asp"
        }
    ]
}

};

oop.inherits(VBScriptHighlightRules, TextHighlightRules);

exports.VBScriptHighlightRules = VBScriptHighlightRules;
});

ace.define("ace/mode/vbscript",["require","exports","module","ace/lib/oop","ace/mode/text","ace/mode/vbscript_highlight_rules"], function(require, exports, module) {
"use strict";

var oop = require("../lib/oop");
var TextMode = require("./text").Mode;
var VBScriptHighlightRules = require("./vbscript_highlight_rules").VBScriptHighlightRules;

var Mode = function() {
    this.HighlightRules = VBScriptHighlightRules;
};
oop.inherits(Mode, TextMode);

(function() {
       
    this.lineCommentStart = ["'", "REM"];
    
    this.$id = "ace/mode/vbscript";
}).call(Mode.prototype);

exports.Mode = Mode;
});
