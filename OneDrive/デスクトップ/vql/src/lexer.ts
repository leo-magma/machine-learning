export type LexToken =
  | { kind: "EOF" }
  | { kind: "STRING"; value: string }
  | { kind: "NUMBER"; value: string }
  | { kind: "IDENT"; value: string }
  | { kind: "FIND" }
  | { kind: "IN" }
  | { kind: "WHERE" }
  | { kind: "ORDER" }
  | { kind: "BY" }
  | { kind: "LIMIT" }
  | { kind: "ASC" }
  | { kind: "DESC" }
  | { kind: "PROJECT" }
  | { kind: "MODULE_KW" }
  | { kind: "SEMANTIC" }
  | { kind: "PROCEDURE" }
  | { kind: "ISSUE" }
  | { kind: "NOT" }
  | { kind: "LIKE" }
  | { kind: "MATCHES" }
  | { kind: "AND" }
  | { kind: "OR" }
  | { kind: "LPAREN" }
  | { kind: "RPAREN" }
  | { kind: "COMMA" }
  | { kind: "EQ" }
  | { kind: "NE" }
  | { kind: "GT" }
  | { kind: "GTE" }
  | { kind: "LT" }
  | { kind: "LTE" }
  | { kind: "TRUE" }
  | { kind: "FALSE" }
  | { kind: "NUL" };

/** Case-insensitive keywords (store lowercased keys only). */
const KEYWORD: Record<string, LexToken["kind"]> = {
  find: "FIND",
  in: "IN",
  where: "WHERE",
  order: "ORDER",
  by: "BY",
  limit: "LIMIT",
  asc: "ASC",
  desc: "DESC",
  project: "PROJECT",
  module: "MODULE_KW",
  semantic: "SEMANTIC",
  procedure: "PROCEDURE",
  issue: "ISSUE",
  not: "NOT",
  like: "LIKE",
  matches: "MATCHES",
  and: "AND",
  or: "OR",
  true: "TRUE",
  false: "FALSE",
  null: "NUL",
};

export class Lexer {
  private i = 0;

  constructor(private readonly input: string) {}

  peekChar(): string {
    return this.input[this.i] ?? "";
  }

  skipWsAndComments(): void {
    for (;;) {
      while (/\s/.test(this.peekChar())) this.i++;
      if (this.peekChar() === "#") {
        while (this.peekChar() !== "" && this.peekChar() !== "\n") this.i++;
        continue;
      }
      if (this.peekChar() === "-" && this.input[this.i + 1] === "-") {
        this.i += 2;
        while (this.peekChar() !== "" && this.peekChar() !== "\n") this.i++;
        continue;
      }
      break;
    }
  }

  private readString(): string {
    if (this.peekChar() !== '"') throw this.err('Expected `"`');
    this.i++;
    let out = "";
    while (this.i < this.input.length) {
      const c = this.input[this.i++];
      if (c === "\\") {
        const n = this.input[this.i++] ?? "";
        if (n === "n") out += "\n";
        else if (n === "r") out += "\r";
        else if (n === "t") out += "\t";
        else if (n === "\\" || n === '"') out += n;
        else out += n;
        continue;
      }
      if (c === '"') return out;
      out += c;
    }
    throw this.err("Unterminated string");
  }

  private readNumber(): string {
    const start = this.i;
    if (this.peekChar() === "-") this.i++;
    while (/[0-9]/.test(this.peekChar())) this.i++;
    if (this.peekChar() === ".") {
      this.i++;
      while (/[0-9]/.test(this.peekChar())) this.i++;
    }
    return this.input.slice(start, this.i);
  }

  private readIdent(): string {
    const start = this.i;
    while (/[a-zA-Z0-9_.]/.test(this.peekChar())) this.i++;
    return this.input.slice(start, this.i);
  }

  next(): LexToken {
    this.skipWsAndComments();
    const c = this.peekChar();
    if (c === "") return { kind: "EOF" };

    if (c === '"') return { kind: "STRING", value: this.readString() };

    if ((c === "-" && /[0-9]/.test(this.input[this.i + 1] ?? "")) || /[0-9]/.test(c))
      return { kind: "NUMBER", value: this.readNumber() };

    if (c === "(") {
      this.i++;
      return { kind: "LPAREN" };
    }
    if (c === ")") {
      this.i++;
      return { kind: "RPAREN" };
    }
    if (c === ",") {
      this.i++;
      return { kind: "COMMA" };
    }

    if (c === "=") {
      if (this.input[this.i + 1] === "=") this.i += 2;
      else this.i++;
      return { kind: "EQ" };
    }
    if (c === "!" && this.input[this.i + 1] === "=") {
      this.i += 2;
      return { kind: "NE" };
    }
    if (c === "<" && this.input[this.i + 1] === "=") {
      this.i += 2;
      return { kind: "LTE" };
    }
    if (c === ">" && this.input[this.i + 1] === "=") {
      this.i += 2;
      return { kind: "GTE" };
    }
    if (c === "<") {
      this.i++;
      return { kind: "LT" };
    }
    if (c === ">") {
      this.i++;
      return { kind: "GT" };
    }

    if (/[a-zA-Z_]/.test(c)) {
      const word = this.readIdent();
      const kw = KEYWORD[word.toLowerCase()];
      if (kw) return { kind: kw } as LexToken;
      return { kind: "IDENT", value: word };
    }

    throw this.err(`Unexpected character '${c}'`);
  }

  private err(msg: string): Error {
    const line = this.input.slice(0, this.i).split("\n").length;
    return new Error(`${msg} at offset ${this.i} (line ~${line})`);
  }
}
