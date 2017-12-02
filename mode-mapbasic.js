ace.define("ace/mode/MapBasic_highlight_rules",["require","exports","module","ace/lib/oop","ace/mode/text_highlight_rules"], function(require, exports, module) {
"use strict";

var oop = require("../lib/oop");
var TextHighlightRules = require("./text_highlight_rules").TextHighlightRules;

var MapBasicHighlightRules = function() {
	//Issues:
	// Currently "punctuation.definition.variable.asp"  Innacurately allows "a@a" as a variable name.
	//CONSTANTS ($REPLACE_CONSTANTS$)
	// There areF many constants in MapBasic and the list goes on and on but all require MapBasic.DEF. Do we include? - Probably
	// Might be best to create on the fly from MapBasic.DEF etc. --- REGEX: "Define\s+\b(\w*)\b" return "$1"
	
	//List of functions/procs/etc from MapBasic Reference using RegEx: "(.*)(function|procedure|statement|clause)"
    var keywordMapper = this.createKeywordMapper({
        "keyword.control.asp":  "For|To|Next|Exit|Do|Loop|While|Wend|If|Then|ElseIf|Case|GoTo|End"
			+ "|End Program|Terminate Application|End MapInfo|OnError|Resume|Function|Sub",
        "storage.type.asp": "Dim|Call|Define|Declare|Redim|Randomize|Type",
        "storage.modifier.asp": "",
        "keyword.operator.asp": "Mod|And|Not|Or|as|Contains|Part|Entire|Within|Partly|Entirely|Intersects",
        "constant.language.asp": "$REPLACE_CONSTANTS$",
        "support.class.asp": "",
        "support.class.collection.asp": "",
        "support.constant.asp": "",
        "support.function.asp": "Handler",
        "support.function.event.asp": "SelChangedHandler|WinClosedHandler|WinChangedHandler|WinFocusChangedHandler"
            + "|RemoteMsgHandler|RemoteQueryHandler|RemoteMapGenHandler|ToolHandler|EndHandler|ForegroundTaskSwitchHandler",   
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
    },"identifier",true);

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
        },
        {
            token: "punctuation.definition.comment.asp",
            regex: "'",	
            next: "comment",
            caseInsensitive: true
        },
        {
            token: "storage.type.asp",
            regex: "OnError GoTo",				
            caseInsensitive: true
        },
        {
            token: "punctuation.definition.string.begin.asp",
            regex: '"',
            next: "string"
        },
//Only useful in seperate system.
//        {
//            token: [
//                "punctuation.definition.variable.asp"
//            ],
//            regex: "[a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*"
//        },
        {
            token: "constant.numeric.asp",
            regex: "-?(?:(?:\\&[Hh][0-9a-fA-F]*)|\\b(?:(?:[0-9]+\\.?[0-9]*)|(?:\\.[0-9]+))(?:(?:e|E)(?:\\+|-)?[0-9]+)?)\\b"
        },
        {
            token: ["entity.name.function.asp"],
            regex: "(?:([a-zA-Z_~\\x80-\\xFF][0-9@&%$#!\\x0C\\x09a-zA-Z_~\\x80-\\xFF]*)(?=\\(\\)?))"
        },
        {
            token: ["keyword.operator.asp"],
            regex: "\\-|\\+|\\*|\\/|\\>|\\<|\\=|\\&(?!h[0-9a-fA-F]*\\b)"
        },
        {
            regex: "\\w+",
            token: keywordMapper
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

oop.inherits(MapBasicHighlightRules, TextHighlightRules);

exports.MapBasicHighlightRules = MapBasicHighlightRules;
});
