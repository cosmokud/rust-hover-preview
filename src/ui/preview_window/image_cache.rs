//! Decoded pictures kept by the version of the file they came from, so a second hover of the
//! same file at the same size costs a lookup rather than a decode.

use super::*;

/// A frame's place in the cache, and when it was last asked for. The stamp is a
/// counter rather than a clock, so the order frames are dropped in cannot be
/// changed by the system clock moving.
pub(super) struct ImageCacheEntry {
    /// Held behind an `Arc` because a frame is the size of the box it is decoded into —
    /// 33 MB at 4K — and both the store and every hit used to copy all of it: a cache that
    /// exists to avoid decoding a file was paying a memcpy per answer instead. One allocation
    /// is made where the pixels are, and the cache and the player share it.
    pub(super) frame: Arc<ImageFrame>,
    pub(super) bytes: usize,
    pub(super) last_used: u64,
}

/// The file's modification time and length: what says a file is not the one that
/// was decoded last time.
#[derive(Clone, PartialEq, Eq, Hash)]
pub(super) struct FileVersion {
    pub(super) modified: Option<SystemTime>,
    pub(super) len: u64,
}

/// What a held frame is only valid for.
///
/// The file and its version are the obvious part. The pixel size is there as
/// well because a frame is stored decoded *and scaled*: the same photo shown at
/// 100% and at fit-to-screen really is different pixels, so only the size that
/// was asked for can be handed back for it.
#[derive(Clone, PartialEq, Eq, Hash)]
pub(super) struct ImageCacheKey {
    pub(super) path: PathBuf,
    pub(super) version: FileVersion,
    pub(super) width: u32,
    pub(super) height: u32,
}

#[derive(Default)]
pub(super) struct ImageCache {
    pub(super) entries: HashMap<ImageCacheKey, ImageCacheEntry>,
    pub(super) bytes: usize,
    pub(super) tick: u64,
}

pub(super) static IMAGE_CACHE: Lazy<Mutex<ImageCache>> =
    Lazy::new(|| Mutex::new(ImageCache::default()));

pub(super) fn file_version(path: &Path) -> FileVersion {
    match std::fs::metadata(path) {
        Ok(metadata) => FileVersion {
            modified: metadata.modified().ok(),
            len: metadata.len(),
        },
        Err(_) => FileVersion {
            modified: None,
            len: 0,
        },
    }
}

/// The memory the cache may hold, read from the configuration each time rather
/// than captured, so an edit to `image_cache_mb` applies without a restart.
pub(super) fn image_cache_limit_bytes() -> usize {
    let megabytes = CONFIG
        .lock()
        .map(|config| sanitize_image_cache_mb(config.image_cache_mb))
        .unwrap_or(DEFAULT_IMAGE_CACHE_MB);

    megabytes as usize * 1024 * 1024
}

/// Drop frames, least recently used first, until the cache fits inside `limit`.
///
/// A limit of zero empties it, which is what makes `image_cache_mb = 0` mean
/// "hold nothing" rather than "hold everything until something else is stored".
///
/// The one thing it is for is memory: the budget is read from the configuration on every call,
/// so a figure typed into the tray applies at the next decode rather than at a restart, and a
/// figure typed down frees what is already held at the moment it is set (`trim_image_cache`)
/// rather than at the next decode that happens to pass through. Which frames go is a policy —
/// *least recently used*, not *largest*, and not *first decoded* — because a cache of decoded
/// frames that evicts by size keeps re-decoding a file the user has moved on from and evicts
/// nothing at all when every frame is the same size.
pub(super) fn image_cache_trim(cache: &mut ImageCache, limit: usize) {
    while cache.bytes > limit {
        // Bound to its own statement so the borrow of `entries` has ended before
        // the frame is removed.
        let oldest = cache
            .entries
            .iter()
            .min_by_key(|(_, entry)| entry.last_used)
            .map(|(key, _)| (*key).clone());

        let Some(oldest) = oldest else {
            break;
        };

        if let Some(dropped) = cache.entries.remove(&oldest) {
            cache.bytes -= dropped.bytes;
        }
    }
}

/// Trim the image cache to the configured size now, which is what the tray asks for
/// when a smaller size is chosen: what is over the new budget is freed at the moment
/// it is set rather than at the next decode that happens to pass through here.
pub(crate) fn trim_image_cache() {
    let limit = image_cache_limit_bytes();
    if let Ok(mut cache) = IMAGE_CACHE.lock() {
        image_cache_trim(&mut cache, limit);
    }
}

/// The frame held for `key`, if the cache still has it.
pub(super) fn image_cache_get(key: &ImageCacheKey) -> Option<Arc<ImageFrame>> {
    let limit = image_cache_limit_bytes();
    let mut cache = IMAGE_CACHE.lock().ok()?;

    image_cache_trim(&mut cache, limit);

    cache.tick += 1;
    let tick = cache.tick;

    let entry = cache.entries.get_mut(key)?;
    entry.last_used = tick;

    Some(entry.frame.clone())
}

/// Hold `frame` for `key`, dropping whatever no longer fits beside it.
pub(super) fn image_cache_put(key: ImageCacheKey, frame: Arc<ImageFrame>) {
    let limit = image_cache_limit_bytes();
    let Ok(mut cache) = IMAGE_CACHE.lock() else {
        return;
    };

    image_cache_trim(&mut cache, limit);

    let bytes = frame.pixels.len();
    // A frame larger than the whole budget would evict everything else and still
    // not fit, so it is simply not held.
    if bytes > limit {
        return;
    }

    cache.tick += 1;
    let tick = cache.tick;

    if let Some(previous) = cache.entries.insert(
        key,
        ImageCacheEntry {
            frame,
            bytes,
            last_used: tick,
        },
    ) {
        cache.bytes -= previous.bytes;
    }
    cache.bytes += bytes;

    image_cache_trim(&mut cache, limit);
}
