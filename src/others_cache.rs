//! The three caches under `Performance → Memory → Others`, measured in the bytes they hold rather
//! than in the number of answers they keep.
//!
//! **What these three have in common is that none of them is a picture.** The budgets beside this
//! one — `Image`, `Animated Image`, and the run a pinned picture is walked into — are stacks of
//! BGRA frames, and what one of those costs is a pixel count. What is here is a parsed text
//! document (its own text, a line index, the lines it has drawn), an archive's listing (the names
//! it walked and the counts beside them), and the one picture the image converter developed for a
//! preview too large for the frame cache (the PNG the engine wrote). Each of those was bounded by
//! a count of its own — twenty-four documents, sixteen listings, one picture — which is a ceiling
//! that says nothing to a hand comparing what the app holds against Task Manager, and that made a
//! sweep of a folder of logs re-parse on the way back while one scanned archive sat in the same
//! sixteen slots as a folder of two-line text files.
//!
//! So the figure is megabytes and the order is least-recently-used, and each cache is asked what
//! its own entries cost rather than how many of them there are. `trim` below is the one place the
//! three are held to one total, because a budget split three ways is three answers to a question
//! the user asked once (see `others_cache_limit_bytes`).
//!
//! **This is not the flush the away timer makes, and the flush is not a replacement for it.** The
//! flush gives all three back at once and for a reason of its own — nobody is at the keyboard to
//! hover anything (see `app::away_flush`) — while this is what bounds them while somebody is.

use crate::config::config::others_cache_limit_bytes;

/// What one held entry of any of the three costs, from the parts that make it up: the text or the
/// names it carries, the index built beside them, and whatever has been built on top of them.
///
/// It is an estimate, and deliberately a generous one. What is being answered is "how much does
/// this sweep of a folder leave behind", and a figure that runs a little high costs one document
/// given back earlier rather than a budget that never bites.
pub(crate) fn held_bytes(text: usize, index: usize, built: usize) -> usize {
    text + index + built
}

/// The bytes a table of `entries` offsets costs, which is the index a parsed document and a
/// listing both carry beside their own text.
pub(crate) fn table_bytes(entries: usize) -> usize {
    entries * std::mem::size_of::<usize>()
}

/// Hold the three caches to the one budget the tray names.
///
/// This is called after an entry has been stored, so what is given back is never the answer that
/// has just arrived: the newest entry is the most recently used by construction, and the loop
/// below takes from the other end of each cache.
///
/// The three are asked in turn, one entry each time round, so that a folder of large text files is
/// bounded without first filling the archive listings beside them and a single developed picture is
/// not held while a hundred documents are given back around it. A round that frees nothing ends the
/// loop, which is what keeps a budget of nothing from spinning over three empty caches.
///
/// **No two of the three are held at once.** Each ask takes that cache's own lock and gives it
/// back before the next is made, which is what keeps a trim from being a place two caches meet.
pub(crate) fn trim() {
    let limit = others_cache_limit_bytes();

    while the_three_are_holding() > limit {
        let freed = drop_one_from_each();

        if freed == 0 {
            return;
        }
    }
}

/// What the three are holding between them, each read under its own lock.
fn the_three_are_holding() -> usize {
    crate::text::text_preview::document_bytes_held()
        + crate::readers::archive_listing::listing_bytes_held()
        + crate::engines::imagemagick_render::developed_bytes_held()
}

/// Give one entry back from each of the three, answering what the three freed between them.
fn drop_one_from_each() -> usize {
    crate::text::text_preview::drop_the_oldest_document().unwrap_or(0)
        + crate::readers::archive_listing::drop_the_oldest_listing().unwrap_or(0)
        + crate::engines::imagemagick_render::drop_the_developed_picture().unwrap_or(0)
}

/// Empty all three, which is what the away timer asks for (see `app::away_flush`).
///
/// Nothing on screen is taken by it: what a window is *drawing* is the frame it painted, and where
/// that frame came from one of these the picture is shared rather than copied (`Arc`), so the
/// window keeps the one it has while the cache gives its own reference back.
pub(crate) fn flush() {
    crate::text::text_preview::clear_documents();
    crate::readers::archive_listing::clear_listings();
    crate::engines::imagemagick_render::clear_developed();
}

#[cfg(test)]
mod tests {

    /// A cache the round-robin below can be handed, which the three real ones are not: each of
    /// them keeps its own lock, so a test of the order does not need a document to be parsed to
    /// answer what it is asked.
    trait OthersCache {
        fn bytes_held(&self) -> usize;
        fn drop_oldest(&mut self) -> Option<usize>;
    }

    /// A stand-in for the three, holding `entries` entries of `each` bytes apiece — which is what
    /// the round-robin above is asked about rather than the caches themselves: their own tests
    /// hold their own locks, and a test of the order does not need a document to be parsed.
    struct Fake {
        entries: usize,
        each: usize,
    }

    impl OthersCache for Fake {
        fn bytes_held(&self) -> usize {
            self.entries * self.each
        }

        fn drop_oldest(&mut self) -> Option<usize> {
            if self.entries == 0 {
                return None;
            }
            self.entries -= 1;
            Some(self.each)
        }
    }

    /// The loop `trim` runs, over any three caches rather than the three: what it is asked about
    /// is the order and the ending, and neither needs a file to be parsed to be answered.
    fn trim_the_three(
        one: &mut dyn OthersCache,
        two: &mut dyn OthersCache,
        three: &mut dyn OthersCache,
        limit: usize,
    ) -> bool {
        loop {
            let total = one.bytes_held() + two.bytes_held() + three.bytes_held();

            if total <= limit {
                return true;
            }

            let freed = one.drop_oldest().unwrap_or(0)
                + two.drop_oldest().unwrap_or(0)
                + three.drop_oldest().unwrap_or(0);

            if freed == 0 {
                return false;
            }
        }
    }

    /// **A cache over its budget gives its oldest entries back until the three of them fit, and one
    /// under it gives nothing back.**
    #[test]
    fn a_cache_over_its_budget_gives_the_oldest_entries_back() {
        let mut documents = Fake {
            entries: 10,
            each: 100,
        };
        let mut listings = Fake {
            entries: 0,
            each: 0,
        };
        let mut developed = Fake {
            entries: 0,
            each: 0,
        };

        assert!(
            trim_the_three(&mut documents, &mut listings, &mut developed, 1000),
            "three caches already inside their budget have nothing to give back"
        );
        assert_eq!(documents.entries, 10);

        trim_the_three(&mut documents, &mut listings, &mut developed, 400);
        assert!(
            documents.bytes_held() <= 400,
            "what is left is inside the budget: {} bytes",
            documents.bytes_held()
        );
    }

    /// **The three are asked in turn**, so a sweep of one kind does not wait for another's cache to
    /// fill: a round that is over its budget gives one entry back from each of them.
    #[test]
    fn the_three_are_asked_in_turn() {
        let mut documents = Fake {
            entries: 10,
            each: 100,
        };
        let mut listings = Fake {
            entries: 10,
            each: 100,
        };
        let mut developed = Fake {
            entries: 1,
            each: 400,
        };

        // A thousand bytes over the budget of two thousand: what goes is one entry from each of
        // the three, and not the ten documents on their own.
        trim_the_three(&mut documents, &mut listings, &mut developed, 2000);
        assert!(documents.entries < 10, "the documents gave one back");
        assert!(listings.entries < 10, "and so did the listings");
        assert!(
            developed.entries == 0,
            "and the developed picture, which is one entry"
        );
    }

    /// **A budget of nothing empties them**, which is what makes `others_cache_mb = 0` mean "hold
    /// nothing between hovers" rather than "hold everything until something else is stored".
    #[test]
    fn a_budget_of_nothing_empties_the_caches() {
        let mut documents = Fake {
            entries: 3,
            each: 64,
        };
        let mut listings = Fake {
            entries: 2,
            each: 64,
        };
        let mut developed = Fake {
            entries: 1,
            each: 64,
        };

        trim_the_three(&mut documents, &mut listings, &mut developed, 0);
        assert_eq!(
            documents.entries, 0,
            "nothing is held at a budget of nothing"
        );
        assert_eq!(listings.entries, 0);
        assert_eq!(developed.entries, 0);
    }

    /// **Three empty caches are not a loop.** A budget of nothing over caches with nothing in them
    /// answers on its first look — there is nothing to give back — rather than asking them again
    /// for ever, which is what the `freed == 0` arm is for.
    #[test]
    fn a_trim_with_nothing_left_to_free_ends() {
        let mut documents = Fake {
            entries: 0,
            each: 0,
        };
        let mut listings = Fake {
            entries: 0,
            each: 0,
        };
        let mut developed = Fake {
            entries: 0,
            each: 0,
        };

        assert!(
            trim_the_three(&mut documents, &mut listings, &mut developed, 0),
            "three empty caches are already inside a budget of nothing"
        );

        // And the same answer for a loop that does have to ask: a budget of nothing over caches
        // that hold something gives everything back and ends.
        let mut documents = Fake {
            entries: 2,
            each: 32,
        };
        assert!(trim_the_three(
            &mut documents,
            &mut listings,
            &mut developed,
            0
        ));
        assert_eq!(documents.entries, 0);
    }
}
