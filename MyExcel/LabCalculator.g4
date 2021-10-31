grammar LabCalculator;


/*
 * Parser Rules
 */

compileUnit : expression EOF;

expression :	//LPAREN expression RPAREN #ParenthesizedExpr
				operatorToken=MAX LPAREN (expression (',' expression)*) RPAREN #MaxExpr
				|operatorToken=MIN LPAREN (expression (',' expression)*) RPAREN #MinExpr
				|LPAREN expression RPAREN #ParenthesizedExpr
				|operatorToken=(ADD | SUBTRACT) expression #UpmExpr
				|expression EXPONENT expression #ExponentialExpr
				|expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
				|expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
				|NUMBER #NumberExpr
				|IDENTIFIER #IdentifierExpr
; 

/*
 * Lexer Rules
 */

NUMBER : INT ('.' INT)?; 
IDENTIFIER : [A-Z]+[1-9][0-9]*;

INT : ('0'..'9')+;

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';
MIN : 'min';
MAX : 'max';

WS : [ \t\r\n] -> channel(HIDDEN);