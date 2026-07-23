/**
 * sheetlang — an Excel-like formula language that compiles to JSON Logic and runs
 * on json-logic-engine.
 *
 * These declarations are hand-written so the public API carries real types rather
 * than the `any` that inference from the JSDoc would produce.
 */

// ---------------------------------------------------------------------------
// Core value and logic types
// ---------------------------------------------------------------------------

/** A value a formula can produce or read. Errors are values, not exceptions. */
export type FormulaValue =
  | number
  | string
  | boolean
  | null
  | Date
  | FormulaError
  | FormulaValue[]

/**
 * A compiled formula: an ordinary JSON Logic tree, and therefore plain
 * serialisable data. Left as `unknown` because its shape is open — every function
 * in the library, and any you add, is a valid node.
 */
export type JsonLogic = unknown

/** The variables a formula may reference. Arbitrary nested data, not just a grid. */
export type Data = Record<string, unknown> | unknown

/** The Excel error tokens the library understands. */
export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#CIRCULAR!'

// ---------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------

export interface CompileOptions {
  /**
   * The function names the compiler will accept. Defaults to the standard library.
   * Pass `null` to skip the check entirely (useful when the engine is extended at
   * runtime), or an explicit set to restrict what user-supplied formulas may call.
   */
  functions?: Iterable<string> | null
  /**
   * The functions to emit in the collapsed `{ FN: x }` unary form. Supply this only
   * when a custom engine declares its own `optimizeUnary` methods.
   */
  unaryFunctions?: Iterable<string>
}

export interface EvaluateOptions extends CompileOptions {
  /** A custom engine, e.g. from {@link createEngine}. Defaults to the built-in one. */
  engine?: Engine
  /**
   * Never generate code — interpret on every call. Required under a strict
   * Content-Security-Policy, where the compiler's use of `eval` is blocked.
   * `evaluate` is always interpreted, so this flag only affects {@link build}.
   */
  interpreted?: boolean
}

export interface EngineOptions {
  /** Functions to add to (or override in) the standard library. */
  functions?: Record<string, unknown>
  /** Cap on nesting depth; a formula deeper than this yields an error. */
  maxDepth?: number
  /** Cap on the length of any array a formula may build. */
  maxArrayLength?: number
  /** Cap on the length of any string a formula may build. */
  maxStringLength?: number
  disableInline?: boolean
  disableInterpretedOptimization?: boolean
  permissive?: boolean
  [option: string]: unknown
}

export interface SheetOptions {
  /** A custom engine. Defaults to the built-in one. */
  engine?: Engine
  /** Recalculate only the affected subgraph on a change. Defaults to `true`. */
  incremental?: boolean
  /** Never compile cells; interpret always. Required under a strict CSP. */
  interpreted?: boolean
}

// ---------------------------------------------------------------------------
// Engine
// ---------------------------------------------------------------------------

/**
 * The subset of json-logic-engine's `LogicEngine` this library relies on. A real
 * engine has more; this is what the public API guarantees.
 */
export interface Engine {
  methods: Record<string, unknown>
  run(logic: JsonLogic, data?: Data, options?: { above?: unknown }): unknown
  build(logic: JsonLogic, options?: { top?: boolean; above?: unknown }): (data?: Data) => unknown
  addMethod(name: string, method: unknown, annotations?: unknown): void
}

/** Create an engine preloaded with the Excel function library, plus your own. */
export function createEngine(options?: EngineOptions): Engine

/** The engine used when a caller does not supply one. */
export const defaultEngine: Engine

// ---------------------------------------------------------------------------
// Compile / evaluate / build
// ---------------------------------------------------------------------------

/** Compile formula text into a JSON Logic tree. Throws {@link FormulaSyntaxError}. */
export function compile(source: string, options?: CompileOptions): JsonLogic

/** Compile and run a formula in one step. Always interpreted, so CSP-safe. */
export function evaluate(source: string, data?: Data, options?: EvaluateOptions): FormulaValue

/**
 * Compile once, then run repeatedly — far faster than {@link evaluate} in a loop.
 * Pass `{ interpreted: true }` to run without code generation under a strict CSP.
 */
export function build(source: string, options?: EvaluateOptions): (data?: Data) => FormulaValue

/** Every function name the compiler accepts, sorted. */
export const functionNames: string[]

/** The function library, as engine methods. */
export const functions: Record<string, unknown>

// ---------------------------------------------------------------------------
// The Sheet
// ---------------------------------------------------------------------------

/**
 * A minimal spreadsheet over the compiler: cells hold a literal or an `=` formula,
 * reads are lazy and memoised, dependencies resolve on demand, and recalculation is
 * incremental. Cycles yield `#CIRCULAR!` rather than hanging.
 */
export class Sheet {
  constructor(cells?: Record<string, unknown>, options?: SheetOptions)

  readonly engine: Engine
  readonly incremental: boolean
  readonly interpreted: boolean

  /** Store a cell. An `=`-prefixed string is compiled now, so errors surface here. */
  set(name: string, value: unknown): this
  /** Store many cells at once. */
  setAll(cells: Record<string, unknown>): this
  /** Remove a cell and dirty its dependents. */
  delete(name: string): this
  /** The evaluated value of a cell; `null` for one that was never set. */
  get(name: string): FormulaValue
  /** The compiled JSON Logic for a formula cell, or `null` for a literal. */
  logicFor(name: string): JsonLogic | null
  /** The cells that read `name`, directly or through a range. */
  dependentsOf(name: string): Set<string>
  /** Every cell, evaluated; errors appear as their token, e.g. `#DIV/0!`. */
  toJSON(): Record<string, FormulaValue | string>
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

/** An Excel error value. Flows through a formula rather than throwing. */
export class FormulaError extends Error {
  constructor(code: ErrorCode | string, detail?: string)
  readonly name: 'FormulaError'
  readonly code: ErrorCode | string
  readonly detail?: string
  toString(): string
  toJSON(): { error: string }
}

/** The location Peggy reports for a syntax error. */
export interface SyntaxLocation {
  start: { offset: number; line: number; column: number }
  end: { offset: number; line: number; column: number }
}

/** Raised by {@link compile} when the source cannot be turned into logic. */
export class FormulaSyntaxError extends Error {
  constructor(message: string, location?: SyntaxLocation)
  readonly name: 'FormulaSyntaxError'
  readonly location?: SyntaxLocation
}

/** The canonical error singletons, keyed by token. */
export const ERROR: Record<ErrorCode, FormulaError>

/** Every error token the grammar accepts as a literal. */
export const ERROR_CODES: ErrorCode[]

/** Whether a value is a {@link FormulaError}. */
export function isError(value: unknown): value is FormulaError

// ---------------------------------------------------------------------------
// Dependency analysis
// ---------------------------------------------------------------------------

/** A rectangle of cells, as `RANGE` precedents are recorded. */
export interface PrecedentRect {
  top: number
  left: number
  bottom: number
  right: number
}

/** What a compiled formula reads, extracted statically from its tree. */
export interface Precedents {
  /** Directly referenced names, e.g. `A1`, `taxRate`. */
  cells: Set<string>
  /** Ranges, kept as rectangles rather than expanded. */
  rects: PrecedentRect[]
  /** True when references could not be resolved statically. */
  dynamic: boolean
}

/** Read a formula's precedents straight off its compiled tree. */
export function precedentsOf(logic: JsonLogic): Precedents

// ---------------------------------------------------------------------------
// Reference and date helpers
// ---------------------------------------------------------------------------

/** Every cell name in the rectangle bounded by two corners, row-major. */
export function expandRange(start: string, end: string): string[][]

/** Parse `A1`/`$A$1` into 1-based coordinates, or `null` if it is not an address. */
export function parseA1(ref: string): { col: number; row: number } | null

/** `A` -> 1, `Z` -> 26, `AA` -> 27. */
export function columnToIndex(letters: string): number

/** 1 -> `A`, 27 -> `AA`. */
export function indexToColumn(index: number): string

/** A `Date` as an Excel serial number (days since 1899-12-30). */
export function dateToSerial(date: Date): number

/** An Excel serial number back to a `Date`. */
export function serialToDate(serial: number): Date
