//! `config.ini` across builds: which shape a file is in, and what it takes to read one that is
//! older than this build — and, the other way round, what a build older than the file must not do.
//!
//! Two numbers are written into the settings section, and they answer different questions.
//!
//! **`config_version`** is the shape the file is in — the shape, not the app. A build writes the
//! number this build is at, and a file carrying a smaller one is read through every step that
//! stands above it (see [`apply_pending`]). A file that carries none is at [`BASE_VERSION`], which
//! is what every file written before this question existed is at.
//!
//! **`min_reader`** is the oldest build that can read the file without losing something of it.
//! A build that writes a value, or a key, that a build before it would drop or break raises it, and
//! every build after it answers the raise by reading that file and never writing it — that is
//! [`Stated::is_ahead_of_this_build`], asked by [`AppConfig::save`].
//!
//! The second one is the half that cannot be shipped later. A build already in the field cannot be
//! taught to keep its hands off a file it does not understand, so the guard is written for the
//! builds after it: each one carrying it protects every file written from then on, and it is what
//! makes a downgrade an inconvenience rather than a loss. Without it an older build reads
//! `pin_key=mouse3,space`, cannot name it, writes its own single-key default over it, and the
//! binding is gone with nothing in the file and nothing on the screen saying why.
//!
//! Neither number is a schema. The shape a file is held to is the table
//! [`SETTING_GROUPS`](super::settings_io::SETTING_GROUPS) names, and what a step does is the step.

use configparser::ini::Ini;

use super::defaults::CONFIG_SECTION;

/// The shape every file written before `config_version` existed is in, and the one a file with no
/// `config_version` of its own is read as.
///
/// Absent is the oldest shape rather than the newest, for two reasons: a file that never carried
/// the number is one an older build wrote, and a file from a build *newer* than this is the case
/// that has to be refused rather than guessed at (see [`apply_pending`]).
pub(crate) const BASE_VERSION: u32 = 1;

/// The shape this build writes, and the number a file is stamped with once it has been read
/// through every step standing above it.
///
/// Raising it is one commit that adds one step to [`STEPS`] and nothing else: a number with no
/// step behind it reads a file as older than it is, and a step with no number is a step that never
/// runs.
pub(crate) const CURRENT_VERSION: u32 = 2;

/// The oldest file shape this build still carries steps for.
///
/// A file below this number is brought up to this number rather than to [`CURRENT_VERSION`], and it
/// is brought up by being given the recommended settings whole rather than one key of them at a
/// time (see [`apply_pending`]) — the keys the steps under the cut name are not the keys a file
/// that old holds, so there is nothing under the cut still worth running.
///
/// What is behind the number is the decision that this app no longer supports files older than a
/// particular release, and its cost is paid in `CHANGELOG.md`: the settings of such a file become
/// the recommended ones. It moves on purpose and never on its own, and it is the only thing that
/// ever deletes a step — an app that never raises it keeps every step it ever wrote.
pub(crate) const MIN_SUPPORTED_VERSION: u32 = BASE_VERSION;

/// The oldest build that can read a file this build writes.
///
/// Raised by the first build that writes a key or a value a build before it would break on — and
/// that is the build which carries this guard, from [`BASE_VERSION`] up. From it onward nothing
/// this app writes is destroyed by anything older than it.
///
/// One number for the whole file rather than one per key. The cost is that a build old enough to
/// lose one setting is treated as old enough to lose them all, which is the honest reading: a build
/// that drops a key of a file it does not understand is not one to trust with the rest of it. What
/// it buys is one number to keep and one line to write, with nothing to hold in step with every
/// release — and the per-key question is [`Superseded`]'s, which is a different one.
pub(crate) const SELF_MIN_READER: u32 = 2;

/// What a file's own numbers say, read out of the settings section before anything else is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Stated {
    /// The shape the file says it is in, and the one it is read through the steps from.
    pub(crate) version: u32,
    /// The oldest build the file says may write it.
    pub(crate) min_reader: u32,
}

impl Stated {
    /// The numbers a file states, or the ones a file stating neither is read as.
    ///
    /// A file holding no `config_version` is at [`BASE_VERSION`] and one holding no `min_reader` is
    /// at [`SELF_MIN_READER`], which is what every file written before either key existed carries
    /// by being silent. A number the file holds that this app cannot read is read as the oldest
    /// shape: a file whose number is unreadable is a file no build wrote.
    pub(crate) fn of(ini: &Ini) -> Self {
        Self {
            version: read_number(ini, CURRENT_VERSION_KEY).unwrap_or(BASE_VERSION),
            min_reader: read_number(ini, MIN_READER_KEY).unwrap_or(SELF_MIN_READER),
        }
    }

    /// Whether this build is too old for the file: it may be read, and it must not be written.
    ///
    /// Asked by the reads that freeze the file (`load`, `reload_from_disk`) on their way out of a
    /// [`Ahead`]; what `save` answers is the flag they leave (`settings_io::AHEAD_OF_US`), since a
    /// write reaches that door from callers that have read no file at all.
    ///
    /// A file a newer build left behind holds keys and values this build knows nothing about, and
    /// a write of this build's own view over it is that build's work gone — silently, on a machine
    /// whose owner sees their settings revert and has no line to read about it.
    pub(crate) fn is_ahead_of_this_build(&self) -> bool {
        self.min_reader > SELF_MIN_READER
    }
}

/// What [`apply_pending`] did to a file, in the terms a caller writes the file back on and a
/// person is told about.
#[derive(Clone, Default)]
pub(crate) struct Applied {
    /// The settings this build brought up to a value of its own, in the order they were reached.
    pub keys: Vec<&'static str>,
    /// Whether the whole file was given the recommended settings rather than the keys of it, which
    /// is what a file older than [`MIN_SUPPORTED_VERSION`] is given.
    pub whole_file: bool,
    /// Whether the two numbers were written, which every file written before this build needs:
    /// neither of them was a key this app wrote.
    pub stamped: bool,
    /// The steps that ran, which is what a person is told about and what `to_ini` writes under the
    /// numbers of now.
    ran: Vec<&'static Step>,
}

impl Applied {
    /// Whether there is anything here for a caller to write.
    pub(crate) fn touched_anything(&self) -> bool {
        self.whole_file || self.stamped || !self.keys.is_empty()
    }

    /// What happened, in the words of the steps that happened — nothing where a file was already
    /// at the shape of this build.
    ///
    /// It is what a person is told rather than a thing written into the file: a setting that moves
    /// between one start and the next with nothing said about it is the kind of change a person
    /// reads as the app forgetting what they chose, which is why every step says why it ran.
    pub(crate) fn says_itself(&self) -> Option<String> {
        if self.whole_file {
            return Some(
                "config.ini is older than any shape this build reads: the recommended settings \
                 are used and the file is written with them"
                    .to_string(),
            );
        }

        let why: Vec<&str> = self.ran.iter().map(|step| step.why).collect();

        (!why.is_empty()).then(|| why.join("; "))
    }
}

/// Read a file through the steps standing above its shape, and stamp it with this build's numbers.
///
/// A file at the shape of this build goes through none of them, which is what makes the second read
/// of a file the app has just written a read with nothing to do — the property the watcher depends
/// on, since a file this app writes is one it must not write again (see `settings_io::load` and
/// `watch_the_configuration`).
///
/// A file carrying a `config_version` above [`CURRENT_VERSION`] is refused rather than read: the
/// shape it is in is one this build has no steps for, and reading it as the newest shape there is
/// is how a file silently loses what a build it cannot name had put there.
pub(crate) fn apply_pending(ini: &mut Ini, stated: Stated) -> Result<Applied, Ahead> {
    if stated.version > CURRENT_VERSION {
        return Err(Ahead::NewerThanThisBuild {
            file: stated.version,
            build: CURRENT_VERSION,
        });
    }

    let mut applied = Applied::default();

    if stated.version < MIN_SUPPORTED_VERSION {
        // The cut. A file older than anything this build still carries a step for does not have
        // the shape the steps above the cut name, so it is not brought up one key at a time: it is
        // given the recommended settings whole, by the caller, and stamped. This is the one place
        // this app replaces a file rather than editing it.
        applied.whole_file = true;
        stamp_numbers(ini);
        applied.stamped = true;
        return Ok(applied);
    }

    for step in STEPS.iter().filter(|step| step.from >= stated.version) {
        applied.keys.extend((step.apply)(ini));
        applied.ran.push(step);
        // The number is raised to what the step leaves behind, so the steps after it are asked
        // about the shape the one before it made rather than about the one on disk.
        ini.set(
            CONFIG_SECTION,
            CURRENT_VERSION_KEY,
            Some(step.to.to_string()),
        );
        applied.stamped = true;
    }

    if !writes_numbers(ini) {
        stamp_numbers(ini);
        applied.stamped = true;
    }

    Ok(applied)
}

/// Whether a file states both numbers this build writes, and so says of itself that there is
/// nothing here left to do.
fn writes_numbers(ini: &Ini) -> bool {
    read_number(ini, CURRENT_VERSION_KEY) == Some(CURRENT_VERSION)
        && read_number(ini, MIN_READER_KEY) == Some(SELF_MIN_READER)
}

/// A file that names a shape newer than the build reading it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Ahead {
    /// `config_version` is above [`CURRENT_VERSION`]: a release went out in the wrong order, or a
    /// file was carried from one that has not. The shape it is in is one this build has no steps
    /// for, and guessing at it is how a file loses what it holds.
    NewerThanThisBuild { file: u32, build: u32 },
}

impl Ahead {
    /// What a person is told, since the alternative is settings that revert for a reason nothing on
    /// the screen gives.
    pub(crate) fn says_itself(&self) -> String {
        match self {
            Self::NewerThanThisBuild { file, build } => format!(
                "config.ini is version {file} and this build reads {build}; its settings are used \
                 as they are and the file is not written"
            ),
        }
    }
}

/// One shape of the file, and what a build changed on the way to the next.
///
/// The step is a shape rather than a release: `from` is a number a file carries, `to` is the number
/// it is left carrying, and `apply` is the whole of the difference between the two.
///
/// It is a function rather than a list of keys because not every change is one — a rename moves a
/// value from a key to another, and a retired value names what `retired_end_of_file` reads out of
/// it — and a chain of steps that could not say so would be a chain that grew the `else if` runs it
/// was meant to replace.
#[derive(Debug)]
struct Step {
    /// The `config_version` a file this step reads carries.
    from: u32,
    /// The `config_version` a file this step leaves behind, and the number the file is stamped with
    /// as the step is run — so the next step asks about the shape the last one left, which is what
    /// makes a chain read one file rather than each of them read the same one.
    to: u32,
    /// The change, whose return is the settings it touched — which is what `differs` reads and what
    /// the changelog is told.
    apply: fn(&mut Ini) -> Vec<&'static str>,
    /// What it did, in one line, for whoever asks why a setting moved.
    why: &'static str,
}

/// Every step this build carries, oldest first — and the first entry is the one this schema was
/// built for.
///
/// **`1 → 2`**: a `pin_key` written by a build that had only one key for the setting gains the
/// middle button beside it. `space` was the default for one build, `mouse3` for the two after it,
/// and `mouse3,space` for this one; a file holding either of the first two names a binding this app
/// shipped and a person did not choose, and a file holding anything else — `space,page_down`,
/// `space,mouse3`, `f7` — is a person's, and is left exactly as it is.
///
/// This is the step that answers the question this chain exists for: a person whose file says
/// `space` because the build they installed shipped `space` ends up with the middle button AND the
/// `space`, without being told to open the file and without a reset row that would take every other
/// setting with it.
static STEPS: &[Step] = &[Step {
    from: 1,
    to: 2,
    apply: bring_up_the_pin_key,
    why: "a pin_key shipped by an older build gains the middle button beside it",
}];

/// The `pin_key` of a file written by a build that had a single-key default, brought up to the
/// pair this build ships.
///
/// The values are the ones this app wrote and no others, which is what makes this narrow: a file
/// holding `space,page_down` or `space,mouse3` is one a person wrote from the documentation, and
/// neither of those is a value any build put there.
fn bring_up_the_pin_key(ini: &mut Ini) -> Vec<&'static str> {
    const SHIPPED_BEFORE: &[&str] = &["mouse3", "space"];
    const OF_TODAY: &str = "mouse3,space";

    match setting(ini, "pin_key").as_deref() {
        Some(value) if SHIPPED_BEFORE.contains(&value) => {
            ini.set(CONFIG_SECTION, "pin_key", Some(OF_TODAY.to_string()));
            vec!["pin_key"]
        }
        _ => Vec::new(),
    }
}

/// Where a file says its own shape is: the shape of the settings section, beside the settings it
/// describes rather than in a section of its own, since a section of this file is what holds a
/// list and this is not a list.
pub(crate) const CURRENT_VERSION_KEY: &str = "config_version";

/// And where it says the oldest build that may write it.
pub(crate) const MIN_READER_KEY: &str = "min_reader";

/// The number a file holds under `key`, or `None` where it holds nothing this app can read.
fn read_number(ini: &Ini, key: &str) -> Option<u32> {
    ini.get(CONFIG_SECTION, key)
        .and_then(|value| value.trim().parse().ok())
}

/// The value a file holds under `key`, trimmed of the spaces a hand or an editor leaves around it.
fn setting(ini: &Ini, key: &str) -> Option<String> {
    ini.get(CONFIG_SECTION, key)
        .map(|value| value.trim().to_string())
}

/// The two numbers, each at the value of now.
fn stamp_numbers(ini: &mut Ini) {
    ini.set(
        CONFIG_SECTION,
        CURRENT_VERSION_KEY,
        Some(CURRENT_VERSION.to_string()),
    );
    ini.set(
        CONFIG_SECTION,
        MIN_READER_KEY,
        Some(SELF_MIN_READER.to_string()),
    );
}
