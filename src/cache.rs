//! Bounded LRU cache and deterministic SipHash-2-4 hashing for persistent keys.

use std::borrow::Borrow;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};

// ------------------------------------------------------------------ BoundedCache

struct Entry<K, V> {
    key: K,
    value: V,
    prev: Option<usize>,
    next: Option<usize>,
}

/// Fixed-capacity in-memory cache with O(1) true LRU eviction.
pub(crate) struct BoundedCache<K, V> {
    capacity: usize,
    index: HashMap<K, usize>,
    entries: Vec<Option<Entry<K, V>>>,
    free: Vec<usize>,
    head: Option<usize>,
    tail: Option<usize>,
}

impl<K: Eq + Hash + Clone, V> BoundedCache<K, V> {
    pub(crate) fn new(capacity: usize) -> Self {
        Self {
            capacity,
            index: HashMap::new(),
            entries: Vec::new(),
            free: Vec::new(),
            head: None,
            tail: None,
        }
    }

    #[allow(dead_code)]
    pub(crate) fn len(&self) -> usize {
        self.index.len()
    }

    #[allow(dead_code)]
    pub(crate) fn is_empty(&self) -> bool {
        self.index.is_empty()
    }

    #[allow(dead_code)]
    pub(crate) fn capacity(&self) -> usize {
        self.capacity
    }

    pub(crate) fn contains_key<Q>(&self, key: &Q) -> bool
    where
        K: Borrow<Q>,
        Q: Hash + Eq + ?Sized,
    {
        self.index.contains_key(key)
    }

    /// Look up `key` without changing recency order.
    #[allow(dead_code)]
    pub(crate) fn peek<Q>(&self, key: &Q) -> Option<&V>
    where
        K: Borrow<Q>,
        Q: Hash + Eq + ?Sized,
    {
        let idx = *self.index.get(key)?;
        self.entries[idx].as_ref().map(|entry| &entry.value)
    }

    /// Look up `key` and promote it to most-recently-used.
    pub(crate) fn get<Q>(&mut self, key: &Q) -> Option<&V>
    where
        K: Borrow<Q>,
        Q: Hash + Eq + ?Sized,
    {
        let idx = *self.index.get(key)?;
        self.move_to_front(idx);
        self.entries[idx].as_ref().map(|entry| &entry.value)
    }

    #[allow(dead_code)]
    pub(crate) fn get_mut<Q>(&mut self, key: &Q) -> Option<&mut V>
    where
        K: Borrow<Q>,
        Q: Hash + Eq + ?Sized,
    {
        let idx = *self.index.get(key)?;
        self.move_to_front(idx);
        self.entries[idx].as_mut().map(|entry| &mut entry.value)
    }

    /// Insert or update `key`, promoting it to MRU and evicting the LRU entry when full.
    pub(crate) fn insert(&mut self, key: K, value: V) -> Option<V> {
        if self.capacity == 0 {
            return None;
        }

        if let Some(&idx) = self.index.get(&key) {
            self.move_to_front(idx);
            let slot = self.entries[idx].as_mut().expect("indexed entry exists");
            return Some(std::mem::replace(&mut slot.value, value));
        }

        if self.index.len() >= self.capacity {
            self.evict_lru();
        }

        let idx = if let Some(reused) = self.free.pop() {
            self.entries[reused] = Some(Entry {
                key: key.clone(),
                value,
                prev: None,
                next: self.head,
            });
            reused
        } else {
            let idx = self.entries.len();
            self.entries.push(Some(Entry {
                key: key.clone(),
                value,
                prev: None,
                next: self.head,
            }));
            idx
        };

        if let Some(old_head) = self.head {
            if let Some(entry) = self.entries[old_head].as_mut() {
                entry.prev = Some(idx);
            }
        } else {
            self.tail = Some(idx);
        }
        self.head = Some(idx);
        self.index.insert(key, idx);
        None
    }

    #[allow(dead_code)]
    pub(crate) fn remove<Q>(&mut self, key: &Q) -> Option<V>
    where
        K: Borrow<Q>,
        Q: Hash + Eq + ?Sized,
    {
        let idx = self.index.remove(key)?;
        self.unlink(idx);
        let entry = self.entries[idx].take()?;
        self.free.push(idx);
        Some(entry.value)
    }

    #[allow(dead_code)]
    pub(crate) fn clear(&mut self) {
        self.index.clear();
        self.entries.clear();
        self.free.clear();
        self.head = None;
        self.tail = None;
    }

    /// Iterate entries from most-recently-used to least-recently-used without changing recency.
    pub(crate) fn iter_mru(&self) -> impl Iterator<Item = (&K, &V)> {
        let mut cursor = self.head;
        std::iter::from_fn(move || {
            let idx = cursor?;
            let entry = self.entries[idx].as_ref()?;
            cursor = entry.next;
            Some((&entry.key, &entry.value))
        })
    }

    /// Give back the least recently used entry, which is the one a byte budget takes first.
    ///
    /// A count-bounded cache takes one whenever it is full, by an `insert` of its own; a cache
    /// bounded by *bytes* is one whose keeper asks in a loop until what is held fits, and it is
    /// the keeper who knows what an entry costs (see `others_cache::trim`). Answers the entry it
    /// took, so the keeper can measure what it freed.
    pub(crate) fn pop_lru(&mut self) -> Option<(K, V)> {
        let tail = self.tail?;
        self.unlink(tail);
        let entry = self.entries[tail].take()?;
        self.index.remove(&entry.key);
        self.free.push(tail);
        Some((entry.key, entry.value))
    }

    fn move_to_front(&mut self, idx: usize) {
        if self.head == Some(idx) {
            return;
        }
        self.unlink(idx);
        let old_head = self.head;
        if let Some(entry) = self.entries[idx].as_mut() {
            entry.prev = None;
            entry.next = old_head;
        }
        if let Some(old_head) = old_head {
            if let Some(entry) = self.entries[old_head].as_mut() {
                entry.prev = Some(idx);
            }
        } else {
            self.tail = Some(idx);
        }
        self.head = Some(idx);
    }

    fn unlink(&mut self, idx: usize) {
        let (prev, next) = match self.entries[idx].as_ref() {
            Some(entry) => (entry.prev, entry.next),
            None => return,
        };
        if let Some(p) = prev {
            if let Some(entry) = self.entries[p].as_mut() {
                entry.next = next;
            }
        } else {
            self.head = next;
        }
        if let Some(n) = next {
            if let Some(entry) = self.entries[n].as_mut() {
                entry.prev = prev;
            }
        } else {
            self.tail = prev;
        }
        if let Some(entry) = self.entries[idx].as_mut() {
            entry.prev = None;
            entry.next = None;
        }
    }

    fn evict_lru(&mut self) {
        let Some(tail_idx) = self.tail else {
            return;
        };
        self.unlink(tail_idx);
        if let Some(entry) = self.entries[tail_idx].take() {
            self.index.remove(&entry.key);
            self.free.push(tail_idx);
        }
    }
}

// ------------------------------------------------------------------ StableHasher (SipHash-2-4)

// Standard SipHash-2-4 test-vector keys; fixed so persistent on-disk keys never drift across Rust releases.
const SIP_K0: u64 = 0x0706_0504_0302_0100;
const SIP_K1: u64 = 0x0f0e_0d0c_0b0a_0908;

/// Deterministic SipHash-2-4 hasher for persistent cache keys.
pub(crate) struct StableHasher {
    v0: u64,
    v1: u64,
    v2: u64,
    v3: u64,
    buf: [u8; 8],
    buf_len: usize,
    total_len: usize,
}

impl StableHasher {
    pub(crate) fn new() -> Self {
        Self::with_keys(SIP_K0, SIP_K1)
    }

    pub(crate) fn with_keys(k0: u64, k1: u64) -> Self {
        Self {
            v0: 0x736f_6d65_7073_6575 ^ k0,
            v1: 0x646f_7261_6e64_6f6d ^ k1,
            v2: 0x6c79_6765_6e65_7261 ^ k0,
            v3: 0x7465_6462_7974_6573 ^ k1,
            buf: [0u8; 8],
            buf_len: 0,
            total_len: 0,
        }
    }

    #[inline]
    fn sip_round(&mut self) {
        self.v0 = self.v0.wrapping_add(self.v1);
        self.v1 = self.v1.rotate_left(13);
        self.v1 ^= self.v0;
        self.v0 = self.v0.rotate_left(32);

        self.v2 = self.v2.wrapping_add(self.v3);
        self.v3 = self.v3.rotate_left(16);
        self.v3 ^= self.v2;

        self.v0 = self.v0.wrapping_add(self.v3);
        self.v3 = self.v3.rotate_left(21);
        self.v3 ^= self.v0;

        self.v2 = self.v2.wrapping_add(self.v1);
        self.v1 = self.v1.rotate_left(17);
        self.v1 ^= self.v2;
        self.v2 = self.v2.rotate_left(32);
    }

    #[inline]
    fn compress(&mut self, m: u64) {
        self.v3 ^= m;
        self.sip_round();
        self.sip_round();
        self.v0 ^= m;
    }
}

impl Default for StableHasher {
    fn default() -> Self {
        Self::new()
    }
}

impl Hasher for StableHasher {
    fn finish(&self) -> u64 {
        let mut s = Self {
            v0: self.v0,
            v1: self.v1,
            v2: self.v2,
            v3: self.v3,
            buf: self.buf,
            buf_len: self.buf_len,
            total_len: self.total_len,
        };
        let mut last = ((s.total_len & 0xff) as u64) << 56;
        for i in 0..s.buf_len {
            last |= (s.buf[i] as u64) << (i * 8);
        }
        s.compress(last);
        s.v2 ^= 0xff;
        s.sip_round();
        s.sip_round();
        s.sip_round();
        s.sip_round();
        s.v0 ^ s.v1 ^ s.v2 ^ s.v3
    }

    fn write(&mut self, mut bytes: &[u8]) {
        self.total_len = self.total_len.wrapping_add(bytes.len());

        if self.buf_len > 0 {
            let need = 8 - self.buf_len;
            if bytes.len() < need {
                self.buf[self.buf_len..self.buf_len + bytes.len()].copy_from_slice(bytes);
                self.buf_len += bytes.len();
                return;
            }
            self.buf[self.buf_len..8].copy_from_slice(&bytes[..need]);
            let m = u64::from_le_bytes(self.buf);
            self.compress(m);
            self.buf_len = 0;
            bytes = &bytes[need..];
        }

        while bytes.len() >= 8 {
            let m = u64::from_le_bytes(bytes[..8].try_into().expect("8 bytes"));
            self.compress(m);
            bytes = &bytes[8..];
        }

        if !bytes.is_empty() {
            self.buf[..bytes.len()].copy_from_slice(bytes);
            self.buf_len = bytes.len();
        }
    }
}

/// Deterministic 64-bit hash of any `Hash` value.
#[allow(dead_code)]
pub(crate) fn stable_hash64<T: Hash + ?Sized>(value: &T) -> u64 {
    let mut hasher = StableHasher::new();
    value.hash(&mut hasher);
    hasher.finish()
}

/// Deterministic 64-bit hash of raw bytes (matches standard SipHash-2-4 test vectors).
#[allow(dead_code)]
pub(crate) fn stable_hash_bytes(bytes: &[u8]) -> u64 {
    let mut hasher = StableHasher::new();
    hasher.write(bytes);
    hasher.finish()
}

/// Deterministic 16-character lowercase hex digest of any `Hash` value.
#[allow(dead_code)]
pub(crate) fn stable_hash_hex<T: Hash + ?Sized>(value: &T) -> String {
    format!("{:016x}", stable_hash64(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bounded_cache_evicts_least_recently_used() {
        let mut cache = BoundedCache::new(3);
        cache.insert("a", 1);
        cache.insert("b", 2);
        cache.insert("c", 3);
        assert_eq!(cache.len(), 3);

        // Touch "a" so "b" becomes LRU.
        assert_eq!(cache.get(&"a"), Some(&1));
        cache.insert("d", 4);

        assert_eq!(cache.len(), 3);
        assert_eq!(cache.peek(&"b"), None);
        assert_eq!(cache.peek(&"a"), Some(&1));
        assert_eq!(cache.peek(&"c"), Some(&3));
        assert_eq!(cache.peek(&"d"), Some(&4));
    }

    #[test]
    fn bounded_cache_update_promotes_and_returns_previous() {
        let mut cache = BoundedCache::new(2);
        assert_eq!(cache.insert("a", 1), None);
        assert_eq!(cache.insert("b", 2), None);
        assert_eq!(cache.insert("a", 10), Some(1));

        // "b" is now LRU and should be evicted next.
        cache.insert("c", 3);
        assert_eq!(cache.peek(&"b"), None);
        assert_eq!(cache.peek(&"a"), Some(&10));
        assert_eq!(cache.peek(&"c"), Some(&3));
    }

    #[test]
    fn bounded_cache_remove_and_zero_capacity() {
        let mut cache = BoundedCache::new(2);
        cache.insert("a", 1);
        cache.insert("b", 2);
        assert_eq!(cache.remove(&"a"), Some(1));
        assert_eq!(cache.len(), 1);
        cache.insert("c", 3);
        cache.insert("d", 4);
        assert_eq!(cache.len(), 2);
        assert_eq!(cache.peek(&"b"), None);
        assert_eq!(cache.peek(&"c"), Some(&3));
        assert_eq!(cache.peek(&"d"), Some(&4));

        let mut zero: BoundedCache<&str, i32> = BoundedCache::new(0);
        zero.insert("x", 1);
        assert!(zero.is_empty());
        assert_eq!(zero.get(&"x"), None);
    }

    #[test]
    fn stable_hasher_matches_official_siphash_2_4_vectors() {
        // Published Aumasson & Bernstein SipHash-2-4 vectors for k = 00..0f, input = 00..(len-1).
        const EXPECTED: [(usize, u64); 6] = [
            (0, 0x726f_db47_dd0e_0e31),
            (1, 0x74f8_39c5_93dc_67fd),
            (7, 0xab02_00f5_8b01_d137),
            (8, 0x93f5_f579_9a93_2462),
            (15, 0xa129_ca61_49be_45e5),
            (16, 0x3f2a_cc7f_57c2_9bdb),
        ];

        for (len, expected) in EXPECTED {
            let input: Vec<u8> = (0..len).map(|i| i as u8).collect();
            assert_eq!(
                stable_hash_bytes(&input),
                expected,
                "SipHash-2-4 vector mismatch at len={len}"
            );

            // Chunked writes must yield the exact same digest as a single write.
            let mut chunked = StableHasher::new();
            for chunk in input.chunks(3) {
                chunked.write(chunk);
            }
            assert_eq!(chunked.finish(), expected, "chunked mismatch at len={len}");
        }
    }

    #[test]
    fn stable_hash_hex_is_deterministic_and_16_hex_chars() {
        let a = stable_hash_hex("c:\\docs\\report.docx");
        let b = stable_hash_hex("c:\\docs\\report.docx");
        let c = stable_hash_hex("c:\\docs\\other.docx");
        assert_eq!(a, b);
        assert_ne!(a, c);
        assert_eq!(a.len(), 16);
        assert!(a.chars().all(|ch| ch.is_ascii_hexdigit()));
    }
}
