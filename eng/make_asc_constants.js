const input = `
    ASC_PIPE = AscW("|")
    ASC_CARET = AscW("^")
    ASC_DOLLAR = AscW("$")
    ASC_QUESTION = AscW("?")
    ASC_STAR = AscW("*")
    ASC_PLUS = AscW("+")
    ASC_LCURLY = AscW("{")
    ASC_0 = AscW("0")
    ASC_1 = AscW("1")
    ASC_9 = AscW("9")
    ASC_COMMA = AscW(",")
    ASC_RCURLY = AscW("}")
    ASC_PERIOD = AscW(".")
    ASC_BACKSLASH = AscW("\\")
    ASC_LPAREN = AscW("(")
    ASC_EQUALS = AscW("=")
    ASC_EXCLAMATION = AscW("!")
    ASC_COLON = AscW(":")
    ASC_RPAREN = AscW(")")
    ASC_LBRACKET = AscW("[")
    ASC_RBRACKET = AscW("]")
    ASC_LC_A = AscW("a")
    ASC_LC_B = AscW("b")
    ASC_LC_C = AscW("c")
    ASC_LC_D = AscW("d")
    ASC_LC_F = AscW("f")
    ASC_LC_N = AscW("n")
    ASC_LC_S = AscW("s")
    ASC_LC_T = AscW("t")
    ASC_LC_R = AscW("r")
    ASC_LC_U = AscW("u")
    ASC_LC_V = AscW("v")
    ASC_LC_W = AscW("w")
    ASC_LC_X = AscW("x")
    ASC_LC_Z = AscW("z")
    ASC_UC_A = AscW("A")
    ASC_UC_B = AscW("B")
    ASC_UC_D = AscW("D")
    ASC_UC_S = AscW("S")
    ASC_UC_W = AscW("W")
    ASC_UC_Z = AscW("Z")
    ASC_MINUS = AscW("-")
`;

const input_lines = input.split("\n").filter(line => !line.match(/^\s*$/));

const output_lines = input_lines
  .map(line => {
    const m = /([A-Za-z0-9_]+).*AscW.\"(.)/.exec(line);  //([A-Z_]+) \= AscW\(\"(.)\"\)/.exec(line);
    const varName = m[1];
    const character = m[2];
    return `Const ${varName} As Long = ${character.charCodeAt(0)}  ' ${character}`;
  });

console.log(output_lines.join("\n"))