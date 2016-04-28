﻿module Hash

/// <summary>
///     Computes the Jenkins One-at-a-time general-purpose hash function.
///     See: https://en.wikipedia.org/wiki/Jenkins_hash_function#one-at-a-time
/// </summary>
/// <param name="key">A sequence of values of generic type.</param>
/// <returns>Returns a 32-bit integer.</returns>
let jenkinsOneAtATimeHash(key: seq<'a>) : int32 =
    // https://en.wikipedia.org/wiki/Jenkins_hash_function#one-at-a-time
    let hash = Seq.fold (fun hash elem ->
                   let hash1 = hash + uint32(elem.GetHashCode())
                   let hash2 = hash1 + (hash1 <<< 10)
                   hash2 ^^^ (hash2 >>> 6)
               ) 0u key
    let hash1 = hash + (hash <<< 3)
    let hash2 = hash1 ^^^ (hash1 >>> 11)
    let hash3 = hash2 + (hash2 <<< 15)
    int32 hash3

let cantorPair(k_1: uint32)(k_2: uint32)(r: uint32) : uint32 =
    (((k_1 + k_2) * (k_1 + k_2 + 1u))/2u + k_2) % r
