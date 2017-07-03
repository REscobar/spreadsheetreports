grammar SpreadSheetGrammar;

compilationUnit: statements;

statements: statements statement |
			statement;

statement:	assignStatement |
			ifStatement |
			LBRACE statements RBRACE;

assignStatement: memberAccessExpression '=' expression STATEMENTEND;

ifStatement:	'if' LPAREN expression RPAREN statement ('else' statement )?;

expression: expression '&&' relationalExpression |
			expression '||' relationalExpression |
			relationalExpression ;

relationalExpression:  relationalExpression '>'  addExpression | 
                       relationalExpression '<'  addExpression | 
                       relationalExpression '<=' addExpression | 
                       relationalExpression '>=' addExpression |
                       relationalExpression '==' addExpression |
                       relationalExpression '!=' addExpression |
                       addExpression;

addExpression: addExpression '+' multExpression |
			   addExpression '-' multExpression |
			   multExpression;

multExpression: multExpression '*' negateExpression |
				multExpression '/' negateExpression |
				negateExpression;

negateExpression: '-' primaryExpression |
				  primaryExpression;

primaryExpression: literalExpression |
				   memberAccessExpression;

memberAccessExpression: memberAccessExpression '.' identifier |
					identifier;


literalExpression: numberLiteral |
				  stringLiteral |
					LPAREN expression RPAREN;

identifier: CHAR+;

numberLiteral: (DIGIT+)('.' DIGIT+)?;

stringLiteral: '"'(CHAR*)'"';

STATEMENTEND : ';';
LPAREN: '(';
RPAREN: ')';
LBRACE: '{';
RBRACE: '}';
CHAR : [a-zA-Z];
DIGIT: [0-9];

WS
   : [\t\r\n ] -> skip
   ;