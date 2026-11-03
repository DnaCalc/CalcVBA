# VBA grammar

Language reference for in-scope VBA behaviour. Public VBA specifications and
Excel/VBA observations decide uncertain cases. See [architecture](../ARCHITECTURE.md)
for current implementation limits; this reference does not claim complete support.

## Notation

This scaffold uses EBNF-style project notation:

- `A = B C` means sequence.
- `A | B` means alternative.
- `A?` means optional.
- `A*` means zero or more.
- `A+` means one or more.
- quoted words are keywords or punctuation in source text, compared case-insensitively for
  keywords.
- ALL_CAPS names are token classes.
- lower_snake_case names are grammar productions.

## Top-Level Grammar Scaffold

```ebnf
source_file        = file_preamble? module_body EOF ;
file_preamble      = (attribute_line | option_line | blank_line | comment_line)* ;
module_body        = module_item* ;
module_item        = option_line
                   | attribute_line
                   | declaration
                   | procedure_decl
                   | property_decl
                   | event_decl
                   | implements_decl
                   | statement_line
                   | blank_line
                   | comment_line ;

option_line        = "Option" option_kind option_value? NEWLINE? ;
option_kind        = "Base" | "Compare" | "Explicit" | "Private" ;
option_value       = expression | identifier ;
attribute_line     = "Attribute" qualified_identifier "=" expression NEWLINE? ;
statement_line     = logical_line ;
logical_line       = line_prefix? statement_list? line_comment? NEWLINE? ;
statement_list     = statement (":" statement)* ;
line_comment       = COMMENT ;
line_prefix        = LABEL ":" | LINE_NUMBER ;
blank_line         = trivia* NEWLINE ;
comment_line       = trivia* COMMENT NEWLINE? ;
trivia             = WHITESPACE | LINE_CONTINUATION ;
```

## Declarations

```ebnf
declaration        = const_decl
                   | dim_decl
                   | variable_decl_stmt
                   | static_decl
                   | declare_decl
                   | enum_decl
                   | type_decl ;

const_decl         = visibility? "Const" const_item ("," const_item)* ;
const_item         = identifier type_ascription? "=" expression ;

dim_decl           = "Dim" variable_decl_list ;
variable_decl_stmt = ("Dim" | visibility) variable_decl_list ;
static_decl        = "Static" variable_decl_list ;

variable_decl_list = variable_decl ("," variable_decl)* ;
variable_decl      = identifier array_rank? type_ascription? ;
array_rank         = "(" bound_list? ")" ;
bound_list         = bound ("," bound)* ;
bound              = expression ("To" expression)? ;
type_ascription    = "As" new_modifier? type_name ;
new_modifier       = "New" ;

enum_decl          = visibility? "Enum" identifier enum_member* "End" "Enum" ;
enum_member        = identifier ("=" expression)? ;

type_decl          = visibility? "Type" identifier field_decl* "End" "Type" ;
field_decl         = identifier array_rank? type_ascription? ;

declare_decl       = visibility? "Declare" ("PtrSafe")? ("Sub" | "Function")
                     identifier lib_clause alias_clause? parameter_list? type_ascription? ;
lib_clause         = "Lib" STRING_LITERAL ;
alias_clause       = "Alias" STRING_LITERAL ;

event_decl         = visibility? "Event" identifier parameter_list? ;
implements_decl    = "Implements" type_name ;
```

## Procedures And Properties

```ebnf
procedure_decl     = sub_decl | function_decl ;
sub_decl           = visibility? "Sub" identifier parameter_list? statement_block "End" "Sub" ;
function_decl      = visibility? "Function" identifier parameter_list? type_ascription?
                     statement_block "End" "Function" ;

property_decl      = property_get | property_let | property_set ;
property_get       = visibility? "Property" "Get" identifier parameter_list? type_ascription?
                     statement_block "End" "Property" ;
property_let       = visibility? "Property" "Let" identifier parameter_list
                     statement_block "End" "Property" ;
property_set       = visibility? "Property" "Set" identifier parameter_list
                     statement_block "End" "Property" ;

parameter_list     = "(" parameter_decl_list? ")" ;
parameter_decl_list = parameter_decl ("," parameter_decl)* ;
parameter_decl     = parameter_modifier* "ParamArray"? identifier array_rank?
                     type_ascription? default_value? ;
parameter_modifier = "Optional" | "ByVal" | "ByRef" ;
default_value      = "=" expression ;

statement_block    = logical_line* ;
visibility         = "Public" | "Private" | "Friend" | "Global" ;
```

## Statements

```ebnf
statement          = assignment_stmt
                   | call_stmt
                   | if_stmt
                   | select_stmt
                   | for_stmt
                   | for_each_stmt
                   | do_loop_stmt
                   | while_wend_stmt
                   | with_stmt
                   | on_error_stmt
                   | resume_stmt
                   | goto_stmt
                   | gosub_stmt
                   | return_stmt
                   | exit_stmt
                   | erase_stmt
                   | redim_stmt
                   | raise_event_stmt
                   | local_decl_stmt
                   | empty_stmt ;

assignment_stmt    = ("Set" | "Let")? lvalue "=" expression ;
call_stmt          = "Call"? callable_expr argument_syntax? ;
if_stmt            = inline_if_stmt | block_if_stmt ;
inline_if_stmt     = "If" expression "Then" statement_list ("Else" statement_list)? ;
block_if_stmt      = "If" expression "Then" logical_line*
                     elseif_clause* else_clause? "End" "If" ;
elseif_clause      = "ElseIf" expression "Then" logical_line* ;
else_clause        = "Else" logical_line* ;

select_stmt        = "Select" "Case" expression logical_line*
                     case_clause* "End" "Select" ;
case_clause        = "Case" case_selector_list logical_line* ;
case_selector_list = "Else" | case_selector ("," case_selector)* ;
case_selector      = expression | expression "To" expression | "Is" case_compare_op expression ;
case_compare_op    = "=" | "<>" | "<" | "<=" | ">" | ">=" ;

for_stmt           = "For" identifier "=" expression "To" expression ("Step" expression)?
                     logical_line* "Next" identifier? ;
for_each_stmt      = "For" "Each" identifier "In" expression logical_line* "Next" identifier? ;
do_loop_stmt       = "Do" loop_condition? logical_line* "Loop" loop_condition? ;
loop_condition     = ("While" | "Until") expression ;
while_wend_stmt    = "While" expression logical_line* "Wend" ;
with_stmt          = "With" expression logical_line* "End" "With" ;

on_error_stmt      = "On" "Error" ("GoTo" (LABEL | LINE_NUMBER | "0") | "Resume" "Next") ;
resume_stmt        = "Resume" ("Next" | LABEL | LINE_NUMBER)? ;
goto_stmt          = "GoTo" (LABEL | LINE_NUMBER) ;
gosub_stmt         = "GoSub" (LABEL | LINE_NUMBER) ;
return_stmt        = "Return" ;
exit_stmt          = "Exit" ("Sub" | "Function" | "Property" | "For" | "Do") ;
erase_stmt         = "Erase" lvalue ("," lvalue)* ;
redim_stmt         = "ReDim" "Preserve"? redim_item ("," redim_item)* ;
redim_item         = lvalue array_rank ;
raise_event_stmt   = "RaiseEvent" identifier argument_syntax? ;
local_decl_stmt    = dim_decl | static_decl ;
empty_stmt         = ;
```

## Expressions

Expression parsing is Pratt/precedence based. The grammar below names the layers for coverage and
fixtures; the parser does not need to implement it as recursive functions per layer.

```ebnf
expression         = comparison_expr ;
comparison_expr    = concat_expr (compare_op concat_expr)* ;
compare_op         = "=" | "<>" | "<" | "<=" | ">" | ">=" | "Is" | "Like" ;
concat_expr        = additive_expr ("&" additive_expr)* ;
additive_expr      = multiplicative_expr (("+" | "-") multiplicative_expr)* ;
multiplicative_expr = unary_expr (("*" | "/" | "\\") unary_expr)* ;
unary_expr         = ("+" | "-" | "Not") unary_expr | exponent_expr ;
exponent_expr      = postfix_expr ("^" unary_expr)? ;
postfix_expr       = primary_expr postfix_part* ;
postfix_part       = member_part | bang_part | index_or_call_part ;
member_part        = "." identifier ;
bang_part          = "!" identifier ;
index_or_call_part = "(" argument_list? ")" ;

primary_expr       = literal
                   | identifier
                   | "New" type_name
                   | "Nothing"
                   | parenthesized_expr
                   | type_of_expr ;
parenthesized_expr = "(" expression ")" ;
type_of_expr       = "TypeOf" expression "Is" type_name ;

argument_syntax    = argument_list | "(" argument_list? ")" ;
argument_list      = argument ("," argument)* ;
argument           = named_argument | expression? ;
named_argument     = identifier ":=" expression ;
callable_expr      = postfix_expr ;
lvalue             = postfix_expr ;
```

## Lexical Token Classes

```ebnf
identifier         = IDENTIFIER | BRACKETED_IDENTIFIER ;
type_name          = qualified_identifier | builtin_type ;
qualified_identifier = identifier ("." identifier)* ;
builtin_type       = "Boolean" | "Byte" | "Currency" | "Date" | "Decimal" | "Double"
                   | "Integer" | "Long" | "LongLong" | "LongPtr" | "Object" | "Single"
                   | "String" | "Variant" ;
literal            = STRING_LITERAL | DATE_LITERAL | NUMERIC_LITERAL | "True" | "False"
                   | "Empty" | "Null" ;
```
