// Type-level helpers of the conformance checks (doc-conformance.ts, stream-conformance.ts): no runtime content.
//
// A check (`Check<I, P, Rules, Except>`) compiles only when every public member of an index.d.ts type `P` conforms to
// the runtime class `I`. On failure the error names the members that do not conform (`name`, `missing:name`,
// `readonly:name`, and `stale:name` for a listed deviation that conforms now).
//
// What is compared, for each member of the declared type:
// - presence: the runtime class has the member (`missing:name` otherwise);
// - optionality: a member declared required is not optional at runtime;
// - readonly: a member that is readonly at runtime (a getter without setter, a `readonly` field) is declared readonly
//   (`readonly:name` otherwise); a member declared readonly may be writable at runtime;
// - the type of a property: the runtime type is assignable to the declared one (what a read gives), recursively through
//   objects, arrays and tuples, promises (`promised`) and async generators, to a depth of {@link CheckDepth} levels (`unknown`
//   beyond); optional / required fields are compared at every level;
// - a method (any function type that the runtime gives): every declared overload must be implemented by some runtime
//   overload: the declared parameters are accepted by the runtime ones (a declared parameter type is assignable to the
//   runtime one, a parameter declared optional is optional at runtime, the declaration has no parameter that the
//   runtime does not take) and the runtime result is assignable to the declared one. Up to 5 overloads are read
//   (TypeScript gives a function with fewer overloads with its first one repeated); a function with more overloads is
//   compared through its last 5. Generic signatures are compared with their type parameters at their constraints;
// - a callback (a function type that the user gives: a parameter of a method, or one of its members): compared the
//   other way round, as TypeScript compares function types: the runtime passes arguments of the declared types (more
//   arguments than declared, not fewer), and the declared result is assignable to the one the runtime expects.
//
// What is not compared:
// - the back-references `worksheet`, `workbook` and `master`: a row or a cell of the streaming API belongs to a stream
//   worksheet, which index.d.ts does not model as a Worksheet; the classes type them as the narrow interfaces they use
//   (`CellWorksheet`, `RowWorksheet`...);
// - members whose name starts with `_` (internal by convention), and runtime members index.d.ts does not declare;
// - the type a setter accepts (`T[K]` is the type a read gives), readonly-ness below the members of the class, and
//   constructors (stream-conformance.ts compares the arguments of the stream constructors);
// - leaf types (dates, buffers, blobs, regular expressions...): compared by TypeScript as they are;
// - beyond {@link CheckDepth} levels of nesting (an array, a promise, a function or an object is a level).
//
// The rules `Rules` of a check (a union) say how some types are compared wherever they appear:
// - a substitution `[declared type, runtime class]` replaces a declared type that is exactly `P` (an intersection such
//   as `Row & { extra: string }` is not) and a runtime type that is exactly `I` by `I`, compared as is: each class of
//   the pairs is checked on its own, so that a check reports only its own members, and the comparison does not walk
//   the same class again in every member that gives one;
// - a {@link ComparedAs} rule replaces a runtime type (exactly that type; a member of a union is replaced on its own)
//   by another one: the known deviations of a type that appears in several members (rule 2 of the alignment: a value
//   that only invalid input or a malformed file gives), each one with its reason where the rules are declared. A rule
//   that no longer has an effect is not reported.
// The members of `Except` (the known deviations of the members of one class, rule 2) are compared with the declarations
// given there instead of the ones of `P`; the check fails with `stale:name` once such a member conforms to `P` itself.

/** The back-references that are not compared (see the file comment). */
export type BackReference = "worksheet" | "workbook" | "master";

/** The depth to which the members are compared (see the file comment). */
export type CheckDepth = 10;

/** Depth counter of {@link Erase}. */
type Previous = [never, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9];

/** `true` when `A` and `B` are the same type (not merely assignable to each other), else `false`. */
export type Equals<A, B> = (<G>() => G extends A ? 1 : 2) extends <G>() => G extends B ? 1 : 2 ? true : false;

/** A rule of a check: the runtime type `Runtime` is compared as `As` (see the file comment). */
export interface ComparedAs<Runtime, As> {
    /** The runtime type replaced. */
    runtime: Runtime;
    /** What it is compared as. */
    as: As;
}

/** The side of a comparison a type comes from: the declaration (index.d.ts) or the runtime class. */
type Side = "declared" | "runtime";

/**
 * Which way a value goes: from the runtime to the user (`"out"`: a property, a result, an argument given to a callback)
 * or from the user to the runtime (`"in"`: an argument, the result of a callback).
 */
type Flow = "out" | "in";

/** Types compared by TypeScript as they are (not walked). */
type Leaf = Date | RegExp | ArrayBuffer | ArrayBufferView | Blob | string | number | boolean | bigint | symbol | null | undefined | void;

/**
 * What replaces `T` (one member of a union) under the rules `R`, else `never`: `["class", I]`, the runtime class of a
 * substitution (on the declared side when `T` is exactly its declared type, on the runtime side when it is exactly the
 * runtime class), compared as is; on the runtime side `["as", A]`, the type a {@link ComparedAs} rule gives, compared as
 * `T` would be. `any` is never replaced.
 */
type Replace<T, R, W extends Side> = 0 extends 1 & T
    ? never
    : R extends readonly [infer P, infer I]
      ? Equals<T, W extends "declared" ? P : I> extends true
          ? ["class", I]
          : never
      : W extends "runtime"
        ? R extends ComparedAs<infer From, infer As>
            ? Equals<T, From> extends true
                ? ["as", As]
                : never
            : never
        : never;

/**
 * The overloads of a function type, as a union of `[parameters, result]` pairs (up to 5 overloads: see the file
 * comment).
 */
export type Overloads<F> = F extends {
    (...args: infer A1): infer R1;
    (...args: infer A2): infer R2;
    (...args: infer A3): infer R3;
    (...args: infer A4): infer R4;
    (...args: infer A5): infer R5;
}
    ? [A1, R1] | [A2, R2] | [A3, R3] | [A4, R4] | [A5, R5]
    : never;

/** The intersection of the members of the union `U`. */
type UnionToIntersection<U> = (U extends unknown ? (k: U) => void : never) extends (k: infer I) => void ? I : never;

/**
 * {@link Erase} of a function type: the intersection of its overloads, compared with strict variance (contravariant
 * parameters, covariant result). A function the runtime gives (`F` "out": a method, a function property) is a function
 * of one parameter, the tuple of its parameters, so that the declared tuple must be assignable to the runtime one: no
 * more parameters, no optional one that the runtime requires. A function the user gives (`F` "in": a callback) is a
 * plain function type: the runtime may pass more arguments than it declares, not fewer. A function with an `any`
 * parameter (`...args: any[]` of the Node.js typings) is a plain function type in both cases.
 */
type EraseFunction<T, D extends number, R, W extends Side, F extends Flow> = UnionToIntersection<
    Overloads<T> extends infer O
        ? O extends [infer A extends readonly unknown[], infer Result]
            ? F extends "out"
                ? 0 extends 1 & A[number]
                    ? (...args: A) => Erase<Result, Previous[D], R, W, F>
                    : (parameters: { [I in keyof A]: Erase<A[I], Previous[D], R, W, "in"> }) => Erase<Result, Previous[D], R, W, F>
                : (...args: { [I in keyof A]: Erase<A[I], Previous[D], R, W, "out"> }) => Erase<Result, Previous[D], R, W, F>
            : never
        : never
>;

/** {@link Erase} of one member of a union that no rule replaces. */
type EraseStructure<T, D extends number, R, W extends Side, F extends Flow> =
    T extends PromiseLike<infer V>
        ? { promised: Erase<V, Previous[D], R, W, F> }
        : T extends AsyncGenerator<infer Y, infer Result>
          ? { yielded: Erase<Y, Previous[D], R, W, F>; returned: Erase<Result, Previous[D], R, W, F> }
          : T extends (...args: never) => unknown
            ? EraseFunction<T, D, R, W, F>
            : T extends Leaf
              ? T
              : T extends readonly (infer E)[]
                ? number extends T["length"]
                    ? Erase<E, Previous[D], R, W, F>[]
                    : { -readonly [I in keyof T]: Erase<T[I], Previous[D], R, W, F> }
                : T extends object
                  ? { [K in keyof T as K extends BackReference | `_${string}` ? never : K]: Erase<T[K], Previous[D], R, W, F> }
                  : T;

/** {@link Erase} of one member of a union. */
type EraseOne<T, D extends number, R, W extends Side, F extends Flow> = 0 extends 1 & T
    ? T
    : Replace<T, R, W> extends infer X
      ? [X] extends [never]
          ? EraseStructure<T, D, R, W, F>
          : X extends ["class", infer I]
            ? I
            : X extends ["as", infer A]
              ? Erase<A, D, R, W, F>
              : never
      : never;

/**
 * The structure of a type as the checks compare it (see the file comment): without back-references nor `_` members,
 * functions as the intersection of their overloads, recursively to depth `D` (`unknown` beyond), the types of the rules
 * `R` replaced. `W` tells which side of the comparison `T` comes from.
 */
export type Erase<T, D extends number, R, W extends Side, F extends Flow = "out"> = [D] extends [never]
    ? unknown
    : T extends unknown
      ? EraseOne<T, D, R, W, F>
      : never;

/** `true` when the member `K` of `T` is readonly (a getter without setter, a `readonly` field). */
type IsReadonly<T, K extends keyof T> = Equals<Pick<T, K>, Readonly<Pick<T, K>>>;

/**
 * The mismatch of the member `K` that the runtime class `I` and the declared type `P` both have: `K` when its type or
 * optionality does not conform, `readonly:K` when it is readonly at runtime only, else `never`.
 */
type MemberMismatch<I, P, K extends keyof I & keyof P, R> = [Erase<Pick<I, K>, CheckDepth, R, "runtime">] extends [
    Erase<Pick<P, K>, CheckDepth, R, "declared">,
]
    ? IsReadonly<I, K> extends true
        ? IsReadonly<P, K> extends true
            ? never
            : `readonly:${K & string}`
        : never
    : K;

/**
 * The members of `P` that the runtime class `I` lacks (`missing:name`) or does not conform to (`name`, `readonly:name`:
 * see the file comment), under the rules `R`.
 */
export type Mismatches<I, P, R = never> = {
    [K in Exclude<keyof P, BackReference>]-?: K extends keyof I ? MemberMismatch<I, P, K, R> : `missing:${K & string}`;
}[Exclude<keyof P, BackReference>];

/** The member name of a mismatch (`readonly:name` and `missing:name` give `name`). */
type MemberName<M> = M extends `readonly:${infer N}` | `missing:${infer N}` ? N : M;

/** The declaration `P` with the members of `Except` declared as there. */
type Relaxed<P, Except> = Omit<P, keyof Except> & Except;

/** The members of `Except` that conform to their declaration in `P` (`stale:name`). */
type Stale<I, P, Except, R> =
    `stale:${Exclude<keyof Except & keyof P, MemberName<Mismatches<I, Pick<P, keyof Except & keyof P>, R>>> & string}`;

/**
 * `true` when the runtime class `I` conforms to the declared type `P` under the rules `R`, its members `Except`
 * compared with the declarations given there (each one a known deviation, see the file comment), and every member of
 * `Except` still deviates; else the names of the mismatches, and `stale:name` for an exception that conforms now
 * (remove it from the list).
 */
export type Check<I, P, R = never, Except extends { [K in keyof P]?: unknown } = Record<never, never>> = [
    Mismatches<I, Relaxed<P, Except>, R> | Stale<I, P, Except, R>,
] extends [never]
    ? true
    : Mismatches<I, Relaxed<P, Except>, R> | Stale<I, P, Except, R>;

/** Compiles only for `true`: a failing check names its mismatches in the error. */
export type Assert<T extends true> = T;
