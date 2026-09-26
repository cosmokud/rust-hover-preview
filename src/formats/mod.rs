pub(crate) mod archive_formats;
pub(crate) mod calibre_formats;
pub(crate) mod codecs;
pub(crate) mod content_type;
pub(crate) mod design_formats;
pub(crate) mod ebook_formats;
pub(crate) mod font_formats;
pub(crate) mod head;
pub(crate) mod image_formats;
pub(crate) mod libre_formats;
pub(crate) mod magick_formats;
pub(crate) mod native_formats;
pub(crate) mod office_formats;
pub(crate) mod peazip_formats;
pub(crate) mod routing;
pub(crate) mod text_formats;
pub(crate) mod vector_formats;
pub(crate) mod video_formats;

#[cfg(test)]
mod tests {
    use std::path::Path;

    /// The questions this layer is allowed to take the configuration's lock in, and the reason
    /// each one may: they are asked from the engine tiers, which hold no configuration of their
    /// own and have none to be handed, so each reads the file's own entry first and takes the
    /// lock for the read of the lists alone.
    ///
    /// Everywhere else in this layer a lock is a lock a caller may already hold — the hook's gate
    /// and the router are both asked with the configuration in hand — and a lock taken twice on
    /// one thread hangs the thread that asked rather than failing, which is a deadlock no test
    /// can wait for. The question that used to take it in here is asked with the lists passed in
    /// instead (see `content_type::of`).
    const ALLOWED: &[(&str, &str)] = &[
        ("calibre_formats.rs", "is_engine_ebook"),
        ("libre_formats.rs", "engine_page_kind"),
        ("magick_formats.rs", "is_engine_picture"),
        ("peazip_formats.rs", "is_engine_archive"),
    ];

    /// Every place the configuration's lock is taken in this layer, named by the file and by the
    /// question around it, and that the set of them is exactly the engine's own four.
    ///
    /// It is a question the compiler cannot answer — the lock is a global behind a function call
    /// like any other — and the answer is what keeps a hover from hanging: the whole reason the
    /// lists are passed into this layer's questions rather than taken by them.
    #[test]
    fn the_configuration_is_locked_only_where_the_engines_ask() {
        let layer = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("src")
            .join("formats");
        let mut found: Vec<(String, String)> = Vec::new();

        for entry in std::fs::read_dir(&layer).expect("the format layer's own folder") {
            let path = entry.expect("a file of the format layer").path();

            if path.extension().and_then(|extension| extension.to_str()) != Some("rs") {
                continue;
            }

            let file = path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            let source = std::fs::read_to_string(&path).expect("a readable source file");

            // The question the last declaration belongs to, which is nothing in a test module:
            // a test's own lock is its own business, and one there fails rather than hangs.
            let mut question: Option<String> = None;

            for line in source.lines() {
                // A lock named in a comment is not a lock, and the rule this test keeps is
                // written about in more than one of them.
                if line.trim_start().starts_with("//") {
                    continue;
                }

                if line.starts_with("mod ") {
                    question = None;
                }

                if let Some(name) = top_level_fn(line) {
                    question = Some(name);
                }

                if line.contains("CONFIG.lock()") {
                    if let Some(question) = &question {
                        found.push((file.clone(), question.clone()));
                    }
                }
            }
        }

        found.sort();
        found.dedup();

        let mut allowed: Vec<(String, String)> = ALLOWED
            .iter()
            .map(|(file, question)| (file.to_string(), question.to_string()))
            .collect();
        allowed.sort();

        assert_eq!(
            found, allowed,
            "the configuration may be taken in this layer only by the questions the engines ask \
             about a file of theirs, each with the file's own entry read before the lock"
        );
    }

    /// The name of the question a line is written in at the top level of its file: a `fn` written
    /// against the left margin, which is where every question in this layer is declared. A line
    /// inside a module — a test module, above all — belongs to no question of the layer's own.
    fn top_level_fn(line: &str) -> Option<String> {
        let rest = line
            .strip_prefix("fn ")
            .or_else(|| line.strip_prefix("pub fn "))
            .or_else(|| line.strip_prefix("pub(crate) fn "))?;
        let name = rest
            .split(|character: char| !character.is_alphanumeric() && character != '_')
            .next()?;

        (!name.is_empty()).then(|| name.to_string())
    }
}
