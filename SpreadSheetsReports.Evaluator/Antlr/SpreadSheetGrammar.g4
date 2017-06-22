grammar SpreadSheetGrammar;

compilationUnit: statements;

statements: statement |
			statements statement;

statement: assign |

;

expression: 
	
   ;

WS
   : [ \t\r\n] -> skip
   ;