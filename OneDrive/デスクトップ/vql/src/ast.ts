export type FindTarget = "semantic" | "procedure" | "module" | "issue";

export type Scope =
  | { kind: "project"; name: string }
  | { kind: "module"; name: string };

export type SortDirection = "ASC" | "DESC";

export type Query = {
  target: FindTarget;
  scope?: Scope;
  where?: Expr;
  orderBy?: { field: string; direction: SortDirection };
  limit?: number;
};

export type Expr =
  | { type: "literal"; value: string | number | boolean | null }
  | { type: "field"; path: string }
  | { type: "call"; name: "changedSince" | "changedBefore"; date: string }
  | { type: "binary"; op: BinaryOp; left: Expr; right: Expr }
  | { type: "unary"; op: "NOT"; expr: Expr }
  | { type: "in"; expr: Expr; values: Array<string | number | boolean | null> }
  | { type: "notIn"; expr: Expr; values: Array<string | number | boolean | null> }
  | { type: "like"; left: Expr; pattern: string };

export type BinaryOp =
  | "=="
  | "!="
  | ">"
  | "<"
  | ">="
  | "<="
  | "AND"
  | "OR";
