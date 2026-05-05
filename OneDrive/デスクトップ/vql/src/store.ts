export type JsonObject = Record<string, unknown>;

export type IrStore = {
  /** Default project name if rows omit origin.project */
  project: string;
  semantics: JsonObject[];
  procedures: JsonObject[];
  modules: JsonObject[];
  issues: JsonObject[];
};
