let sp = [' ' '\t']+
let id = ['A'-'Z' 'a'-'z' '_'] ['A'-'Z' 'a'-'z' '0'-'9' '_']*
let nl = '\r'? '\n'
let block = "Sub" | "Function" | ("Property " ("Get" | "Let" | "Set"))
let compare = "Binary" | "Text" | "Database"
let scope = "Private" | "Public"
let comment = "'" [^'\r' '\n']* nl as c
let con = ' ' comment | nl
let n = ['0'-'9']+
(* FIXME handle : correctly *)
let sym = ['=' '+' '-' '*' '/' '^' '&' '.' '(' ',' ')' ':' '<' '>']

rule string = parse
| [^'"']+ as x { print_string x; string lexbuf }
| "\"\"" as x { print_string x; string lexbuf }
| '"' as x { print_char x; compile lexbuf }

and decl scope = parse
| (id as x) '$'? {
	print_endline @@ scope ^ " " ^ x;
	decl scope lexbuf
}
| (id as x) " As String * " (n as n) {
	print_endline @@ scope ^ " " ^ x;
	print_endline @@ x ^ " = \"" ^ String.make (int_of_string n) ' ' ^ "\"";
	decl scope lexbuf
}
| (id as x) " As " ("New " as n)? (id as t) {
	(* TODO handle arrays *)
	let pre, init =
		match n, t with
		| None, "Byte" -> "", "CByte(0)"
		| None, "Integer" -> "", "0"
		| None, "Single" -> "", "CSng(0.0)"
		| None, "Long" -> "", "CLng(0)"
		| None, "Double" -> "", "0.0"
		(* decimal *)
		| None, "Date" -> "", "CDate(0)"
		| None, "Currency" -> "", "CCur(0)"
		| None, "String" -> "", "\"\""
		| None, "Object" -> "Set ", "Nothing"
		| None, "Boolean" -> "", "False"
		| None, "Variant" -> "", "Empty"
		| None, _ -> "Set ", "Nothing"
		| Some _, cl -> "Set ", "New " ^ cl
	in
	print_endline @@ scope ^ " " ^ x ^ " 'As " ^ t;
	print_endline @@ pre ^ x ^ " = " ^ init;
	decl scope lexbuf
}
| ", " { decl scope lexbuf }
| con {
	Lexing.new_line lexbuf;
	Option.iter print_string c;
	print_newline ();
	compile lexbuf
}

and args = parse
| (id as x) ('$' | " As " id)? { print_string x; args lexbuf }
| ", " as x { print_string x; args lexbuf }
| ')' (" As " id)? con {
	Lexing.new_line lexbuf;
	print_endline ")";
	Option.iter print_string c;
	compile lexbuf
}

and compile = parse
| comment {
	Lexing.new_line lexbuf;
	print_string c;
	compile lexbuf
}
| ("Implements " id as x) con {
	Lexing.new_line lexbuf;
	print_endline @@ "'" ^ x;
	Option.iter print_string c;
	compile lexbuf
}
| "Next " (id as x) con {
	Lexing.new_line lexbuf;
	print_endline @@ "Next '" ^ x;
	Option.iter print_string c;
	compile lexbuf
}
| ("On Error GoTo 0" as x) con {
	Lexing.new_line lexbuf;
	print_endline x;
	Option.iter print_string c;
	compile lexbuf
}
| "GoTo" { failwith "GoTo aren't supported in VBScript" }
| '"' as x { print_char x; string lexbuf }
| sp | id | sym | n as x { print_string x; compile lexbuf }
| (id as x) '$' { print_string x; compile lexbuf }
| ("Dim" | scope as s) ' ' { decl s lexbuf }
| (scope ' ')? block ' ' id '(' as x { print_string x; args lexbuf }
| ((scope ' ')? block ' ' id as x) "$(" { print_string @@ x ^ "("; args lexbuf }
| nl {
	Lexing.new_line lexbuf;
	print_newline ();
	compile lexbuf
}
| eof { () }


{
let () = compile @@ Lexing.from_channel stdin
}
