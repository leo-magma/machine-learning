import type { BinaryOp, Expr, FindTarget, Query, Scope, SortDirection } from "./ast.js";
import { Lexer, type LexToken } from "./lexer.js";

export function parseQuery(source: string): Query {
  const p = new Parser(source);
  return p.parseQuery();
}

class Parser {
  private lex: Lexer;
  private cur: LexToken;

  constructor(source: string) {
    this.lex = new Lexer(source);
    this.cur = this.lex.next();
  }

  /** Avoid stale discriminant narrowing on `this.cur` after `bump()`. */
  private wideCur(): LexToken {
    return this.cur as unknown as LexToken;
  }

  private bump(): void {
    this.cur = this.lex.next();
  }

  private expect<K extends LexToken["kind"]>(kind: K): Extract<LexToken, { kind: K }> {
    if (this.cur.kind !== kind) throw new Error(`Expected ${kind}, got ${this.cur.kind}`);
    const t = this.cur as Extract<LexToken, { kind: K }>;
    this.bump();
    return t;
  }

  parseQuery(): Query {
    this.expect("FIND");
    const target = this.parseTarget();

    let scope: Scope | undefined;
    if (this.wideCur().kind === "IN") {
      this.bump();
      const scopeKind = this.wideCur().kind;
      if (scopeKind === "PROJECT") {
        this.bump();
        const name = this.expect("STRING").value;
        scope = { kind: "project", name };
      } else if (scopeKind === "MODULE_KW") {
        this.bump();
        const name = this.expect("STRING").value;
        scope = { kind: "module", name };
      } else {
        throw new Error('IN must be followed by project "..." or module "..."');
      }
    }

    let where: Expr | undefined;
    if (this.cur.kind === "WHERE") {
      this.bump();
      where = this.parseExpr();
    }

    let orderBy: Query["orderBy"];
    if (this.cur.kind === "ORDER") {
      this.bump();
      this.expect("BY");
      const field = this.parseFieldPath();
      let direction: SortDirection = "ASC";
      const dirTok = this.wideCur().kind;
      if (dirTok === "ASC") {
        this.bump();
      } else if (dirTok === "DESC") {
        this.bump();
        direction = "DESC";
      }
      orderBy = { field, direction };
    }

    let limit: number | undefined;
    if (this.cur.kind === "LIMIT") {
      this.bump();
      const n = Number(this.expect("NUMBER").value);
      if (!Number.isFinite(n) || n < 0 || !Number.isInteger(n))
        throw new Error("LIMIT must be a non-negative integer");
      limit = n;
    }

    if (this.cur.kind !== "EOF") throw new Error(`Unexpected token ${this.cur.kind} after query`);

    return { target, scope, where, orderBy, limit };
  }

  private parseTarget(): FindTarget {
    switch (this.cur.kind) {
      case "SEMANTIC":
        this.bump();
        return "semantic";
      case "PROCEDURE":
        this.bump();
        return "procedure";
      case "MODULE_KW":
        this.bump();
        return "module";
      case "ISSUE":
        this.bump();
        return "issue";
      default:
        throw new Error(`FIND target must be semantic, procedure, module, or issue`);
    }
  }

  private parseExpr(): Expr {
    return this.parseOr();
  }

  private parseOr(): Expr {
    let left = this.parseAnd();
    while (this.cur.kind === "OR") {
      this.bump();
      const right = this.parseAnd();
      left = { type: "binary", op: "OR", left, right };
    }
    return left;
  }

  private parseAnd(): Expr {
    let left = this.parseUnary();
    while (this.cur.kind === "AND") {
      this.bump();
      const right = this.parseUnary();
      left = { type: "binary", op: "AND", left, right };
    }
    return left;
  }

  private parseUnary(): Expr {
    if (this.cur.kind === "NOT") {
      this.bump();
      return { type: "unary", op: "NOT", expr: this.parseUnary() };
    }
    return this.parseComparison();
  }

  private parseComparison(): Expr {
    const left = this.parsePrimary();

    if (this.wideCur().kind === "NOT") {
      this.bump();
      const afterNot = this.wideCur().kind;
      if (afterNot !== "IN") throw new Error("NOT must be followed by IN");
      this.bump();
      const values = this.parseInList();
      return { type: "notIn", expr: left, values };
    }

    if (this.cur.kind === "IN") {
      this.bump();
      const values = this.parseInList();
      return { type: "in", expr: left, values };
    }

    if (this.cur.kind === "LIKE" || this.cur.kind === "MATCHES") {
      this.bump();
      const pat = this.expect("STRING").value;
      return { type: "like", left, pattern: pat };
    }

    const op = this.parseCmpOp();
    if (!op) return left;

    const right = this.parsePrimary();
    return { type: "binary", op, left, right };
  }

  private parseCmpOp(): BinaryOp | null {
    switch (this.cur.kind) {
      case "EQ":
        this.bump();
        return "==";
      case "NE":
        this.bump();
        return "!=";
      case "GT":
        this.bump();
        return ">";
      case "GTE":
        this.bump();
        return ">=";
      case "LT":
        this.bump();
        return "<";
      case "LTE":
        this.bump();
        return "<=";
      default:
        return null;
    }
  }

  private parseInList(): Array<string | number | boolean | null> {
    this.expect("LPAREN");
    const values: Array<string | number | boolean | null> = [];
    if (this.cur.kind === "RPAREN") {
      this.bump();
      return values;
    }
    for (;;) {
      values.push(this.parseLiteralValue());
      if (this.cur.kind === "COMMA") {
        this.bump();
        continue;
      }
      this.expect("RPAREN");
      return values;
    }
  }

  private parseLiteralValue(): string | number | boolean | null {
    if (this.cur.kind === "STRING") {
      const v = this.cur.value;
      this.bump();
      return v;
    }
    if (this.cur.kind === "NUMBER") {
      const v = Number(this.cur.value);
      this.bump();
      return v;
    }
    if (this.cur.kind === "TRUE") {
      this.bump();
      return true;
    }
    if (this.cur.kind === "FALSE") {
      this.bump();
      return false;
    }
    if (this.cur.kind === "NUL") {
      this.bump();
      return null;
    }
    throw new Error("Expected literal in IN list");
  }

  private parsePrimary(): Expr {
    if (this.cur.kind === "LPAREN") {
      this.bump();
      const e = this.parseExpr();
      this.expect("RPAREN");
      return e;
    }

    if (this.cur.kind === "STRING" || this.cur.kind === "NUMBER") {
      const lit = this.parseAtomicLiteral();
      return { type: "literal", value: lit };
    }

    if (this.cur.kind === "TRUE") {
      this.bump();
      return { type: "literal", value: true };
    }
    if (this.cur.kind === "FALSE") {
      this.bump();
      return { type: "literal", value: false };
    }
    if (this.cur.kind === "NUL") {
      this.bump();
      return { type: "literal", value: null };
    }

    if (this.cur.kind === "IDENT") {
      const id = this.cur.value;
      this.bump();
      const nextAfterIdent = this.wideCur().kind;
      if (nextAfterIdent === "LPAREN") {
        return this.parseCall(id);
      }
      return { type: "field", path: id };
    }

    if (this.cur.kind === "PROJECT") {
      this.bump();
      return { type: "field", path: "project" };
    }

    if (this.cur.kind === "MODULE_KW") {
      this.bump();
      return { type: "field", path: "module" };
    }

    throw new Error(`Unexpected token ${this.cur.kind} in expression`);
  }

  private parseCall(name: string): Expr {
    this.expect("LPAREN");
    if (name === "changedSince" || name === "changedBefore") {
      const d = this.expect("STRING").value;
      this.expect("RPAREN");
      return { type: "call", name, date: d };
    }
    throw new Error(`Unknown function '${name}'`);
  }

  private parseFieldPath(): string {
    const e = this.parsePrimary();
    if (e.type !== "field") throw new Error("ORDER BY expects a field path");
    return e.path;
  }

  private parseAtomicLiteral(): string | number | boolean | null {
    switch (this.cur.kind) {
      case "STRING": {
        const v = this.cur.value;
        this.bump();
        return v;
      }
      case "NUMBER": {
        const v = Number(this.cur.value);
        this.bump();
        return v;
      }
      case "TRUE":
        this.bump();
        return true;
      case "FALSE":
        this.bump();
        return false;
      case "NUL":
        this.bump();
        return null;
      default:
        throw new Error("Expected literal");
    }
  }
}
