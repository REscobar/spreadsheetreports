grammar SpreadSheetGrammar;

compilationUnit: statements;

statements: statement |
			statements statement;

statement: assign
;

assign: expression '=' expression;

expression: 
	memberAccessExpression |
	literalExpression;

memberAccessExpression: memberAccessExpression '.' identifier |
						identifier;

literalExpression: numberLiteral |
				  stringLiteral ;

identifier: CHAR;

numberLiteral: DIGIT+;

stringLiteral: '"'(CHAR*)'"';


CHAR : [a-zA-Z];
DIGIT: [0-9];

WS
   : [\t\r\n] -> skip
   ;