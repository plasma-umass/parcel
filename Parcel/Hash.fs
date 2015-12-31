module Hash

/// <summary>
///     Computes the Jenkins One-at-a-time general-purpose hash function.
///     See: https://en.wikipedia.org/wiki/Jenkins_hash_function#one-at-a-time
/// </summary>
/// <param name="key">A sequence of values of generic type.</param>
/// <returns>Returns a 32-bit integer.</returns>
let jenkinsOneAtATimeHash(key: seq<'a>) : uint32 =
    // https://en.wikipedia.org/wiki/Jenkins_hash_function#one-at-a-time
    let hash = Seq.fold (fun hash elem ->
                   let hash1 = hash + uint32(elem.GetHashCode())
                   let hash2 = hash1 + (hash1 <<< 10)
                   hash2 ^^^ (hash2 >>> 6)
               ) 0u key
    let hash1 = hash + (hash <<< 3)
    let hash2 = hash1 ^^^ (hash1 >>> 11)
    hash2 + (hash2 <<< 15)
