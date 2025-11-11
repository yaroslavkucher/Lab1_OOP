grammar Sheets;

expression
    : expression '*' expression                # MultiplyExpr
    | expression '/' expression                # DivideExpr
    | expression '+' expression                # AddExpr
    | expression '-' expression                # SubtractExpr
    | 'dec(' expression ')'                    # DecrementExpr
    | 'inc(' expression ')'                    # IncrementExpr
    | '-' expression                           # NegateExpr
    | '+' expression                           # AffirmExpr
    | '(' expression ')'                       # ParenExpr
    | expression '=' expression                # EqualExpr
    | expression '<' expression                # LessExpr
    | expression '>' expression                # GreaterExpr
    | 'not(' expression ')'                    # DenialExpr
    | '=' expression                           # EqualSignExpr
    | cell                                     # CellExpr
    | number                                   # NumberExpr
    ;

number      : DIGIT+(('.'|',')DIGIT+)?;
cell        : LETTER+DIGIT+;
 
DIGIT       : [0-9];
LETTER      : [A-Z];

WS          : [ \t\r\n]+ -> skip ; 
