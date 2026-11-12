// Type-level helpers of the conformance checks (doc-conformance.ts, stream-conformance.ts): no runtime content.
//
// A check compiles only when every public member of an index.d.ts type, except the ones listed, conforms to the runtime
// class (its type is assignable to the declared one). On failure the error names the members that do not conform. The
// comparison goes through {@link Erase}:
// - the back-references `worksheet`, `workbook` and `master` are not compared: a row or a cell of the streaming API
//   belongs to a stream worksheet, which index.d.ts does not model as a Worksheet; the classes type them as the narrow
//   interfaces they use (`CellWorksheet`, `RowWorksheet`...);
// - members whose name starts with `_` (internal by convention) are dropped;
// - methods are compared as TypeScript compares methods (bivariant parameters), to a depth of 3 levels;
// - optionally, declared types that are checked on their own are replaced by their runtime class wherever they appear
//   (substitutions `S`: a union of `[declared type, runtime class]` pairs), so that a check reports only its own members.

/** The back-references that are not compared (see the file comment). */
export type BackReference = "worksheet" | "workbook" | "master";

/** Depth counter of {@link Erase}. */
type Previous = [never, 0, 1, 2, 3];

/** A function type compared as a method (bivariant parameters). */
type Method<A extends readonly unknown[], R> = { method(...args: A): R }["method"];

/** The runtime class that replaces `T` (a declared type of a pair of `S`), else `never`; `any` is never replaced. */
type Substitute<T, S> = 0 extends 1 & T ? never : S extends readonly [infer P, infer I] ? ([T] extends [P] ? I : never) : never;

/** {@link Erase} of one member of a union. */
type EraseOne<T, D extends number, S> = [Substitute<T, S>] extends [never]
    ? T extends PromiseLike<infer V>
        ? { promised: Erase<V, Previous[D], S> }
        : T extends AsyncGenerator<infer Y, infer R>
          ? { yielded: Erase<Y, Previous[D], S>; returned: Erase<R, Previous[D], S> }
          : T extends (...args: infer A) => infer R
            ? Method<{ [I in keyof A]: Erase<A[I], Previous[D], S> }, Erase<R, Previous[D], S>>
            : T extends Date | string | number | boolean | null | undefined
              ? T
              : T extends readonly (infer E)[]
                ? Erase<E, Previous[D], S>[]
                : T extends object
                  ? { [K in keyof T as K extends BackReference | `_${string}` ? never : K]: Erase<T[K], Previous[D], S> }
                  : T
    : EraseOne<Substitute<T, S>, D, never>;

/**
 * The structure of a type as the checks compare it: without back-references nor `_` members, recursively to depth
 * `D` (`unknown` beyond), the declared types of the substitutions `S` replaced by their runtime class.
 */
export type Erase<T, D extends number = 3, S = never> = [D] extends [never] ? unknown : T extends unknown ? EraseOne<T, D, S> : never;

/** The members of `P` that the runtime class `I` lacks (`missing:name`) or does not conform to. */
export type Mismatches<I, P, S = never> = {
    [K in Exclude<keyof P, BackReference>]-?: K extends keyof I
        ? [Erase<I[K], 3, S>] extends [Erase<P[K], 3, S>]
            ? never
            : K
        : `missing:${K & string}`;
}[Exclude<keyof P, BackReference>];

/**
 * `true` when `I` conforms to `P` but for the members `Except`, with the substitutions `S`; else the names of the other
 * mismatches.
 */
export type Check<I, P, Except extends keyof P = never, S = never> = [Mismatches<I, Omit<P, Except>, S>] extends [never]
    ? true
    : Mismatches<I, Omit<P, Except>, S>;

/** Compiles only for `true`: a failing check names its mismatches in the error. */
export type Assert<T extends true> = T;
