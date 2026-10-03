use super::*;

/// The one sanitiser, read by four rules: what each of them keeps is the whole of what sixteen
/// lists that are one function are sixteen lists.
///
/// The failures are the ones the rules exist for — a compound name the archive list claims by
/// silently ceasing to be claimed, a C# project of the text list not being a text file, and a
/// dot file the name list cannot find because its dot was read as a separator.
#[test]
fn an_entry_is_read_by_the_rule_its_list_names() {
    assert_eq!(
        Entries::Bare.sanitize(".ZIP, ttf ,nonsense*,,woff2"),
        vec!["zip", "ttf", "woff2"],
        "a typed extension is what a user means, and an entry with a character in it that \
         cannot be one is dropped rather than matched against"
    );

    assert_eq!(
        Entries::Bare.sanitize("tar.gz"),
        Vec::<String>::new(),
        "a bare list has no compound name to keep, and drops the entry rather than splitting \
         one name into two extensions — which is the archive list's whole difference"
    );

    assert_eq!(
        Entries::Compound.sanitize(".ZIP, tar.gz ,nonsense*,,docx"),
        vec!["zip", "tar.gz", "docx"],
        "the archive list keeps the compound a tarball is claimed by"
    );

    assert_eq!(
        Entries::Hashed.sanitize("cs,cshtml,C#"),
        vec!["cs", "cshtml", "c#"],
        "the text list admits a hash so a C# project is a text file"
    );

    assert_eq!(
        Entries::Bare.sanitize("cs,C#"),
        vec!["cs"],
        "and no other list does, so a hash is a stray entry there"
    );

    assert_eq!(
        Entries::Name.sanitize(".gitignore,Makefile,cmakelists.txt,.env.example"),
        vec!["gitignore", "makefile", "cmakelists.txt", "env.example"],
        "a name list reads the dot as part of the name, at the front of one and inside another"
    );
}

/// A list is a set, so the same name typed twice is one name: two entries would be two matches
/// where one file already satisfies the first.
#[test]
fn a_name_typed_twice_is_one_name() {
    assert_eq!(
        Entries::Bare.sanitize("TTF, ttf ,ttf"),
        vec!["ttf"],
        "the case an entry was typed in is not part of the name it is looked up by"
    );
}

/// The whole of the repair: a `config.ini` an older build wrote must come up to the list of now,
/// for every older list this app ever shipped, and the entries must not drift.
///
/// This is walked from the table rather than written out list by list, which is what it is for:
/// an older list no row names is a format that never reaches an installation which already
/// exists, and nothing else in the tree would say so.
#[test]
fn every_older_list_this_app_shipped_is_brought_up_to_the_list_of_now() {
    let mut seen = 0;

    for list in LISTS.iter().filter(|list| list.repaired) {
        let canonical = list.built_in();
        assert!(
            !canonical.is_empty(),
            "a list holds nothing to bring a file up to"
        );

        for older in list.before {
            let mut ini = Ini::new();
            ini.set(list.section, list.key, Some(older.to_string()));

            assert!(
                repair_older_lists(&mut ini),
                "a file holding the older `[{}]` list was not brought up",
                list.section
            );
            assert_eq!(
                ini.get(list.section, list.key).as_deref(),
                Some(canonical.join(",").as_str()),
                "a file holding the older `[{}]` list is not the list of now",
                list.section
            );
            seen += 1;
        }
    }

    assert!(
        seen > 0,
        "no row is repaired, so no list this app changed is brought up"
    );
}

/// A list a person edited is theirs, and the repair is not what decides otherwise: a file
/// holding the built-in entries and one of their own is kept exactly as it is, because the
/// repair can only tell an edit from a list this app wrote by the entries, and one entry the
/// app never shipped is an edit.
#[test]
fn a_list_a_person_edited_is_theirs_and_is_left_exactly_as_it_is() {
    for list in LISTS.iter().filter(|list| list.repaired) {
        let mut ini = Ini::new();
        let edited = format!("{},mine", list.defaults);
        ini.set(list.section, list.key, Some(edited.clone()));

        assert!(
            !repair_older_lists(&mut ini),
            "`[{}]` holds a name this app never shipped and the repair rewrote it",
            list.section
        );
        assert_eq!(
            ini.get(list.section, list.key).as_deref(),
            Some(edited.as_str())
        );
    }
}

/// The six lists that arrived with their kind are not walked, and this is the whole of what that
/// buys: a file holding one of them has been edited by a person, and a person's order is theirs.
///
/// It is also the one thing here that is a fact about a row rather than about a list, and the
/// flag on the row is the only place that says so — so it is held here against a row being given
/// the flag by accident.
#[test]
fn a_section_that_arrived_with_its_kind_is_not_walked() {
    for list in LISTS.iter().filter(|list| !list.repaired) {
        assert!(
            list.before.is_empty(),
            "`[{}]` is not walked and has an older list, so nothing ever reads it",
            list.section
        );
    }

    for name in ["font", "audio", "office", "ebook", "text"] {
        assert!(
            LISTS
                .iter()
                .any(|list| list.section == name && !list.repaired),
            "a kind that arrived with its section is not walked by the repair"
        );
    }
}

/// A list is written and read back as the same entries, which is what a save followed by a load
/// has to be for the settings reset to leave a file's own lists alone.
#[test]
fn a_list_written_out_of_the_configuration_is_the_list_read_back_from_it() {
    let mut config = AppConfig::default();
    for list in LISTS {
        let mut list_as_held = list.built_in();
        list_as_held.push("rhn-test".to_string());
        (list.set)(&mut config, list_as_held);
    }

    let mut ini = Ini::new();
    write_all(&config, &mut ini);

    let mut read_back = AppConfig::default();
    read_all(&ini, &mut read_back);

    for list in LISTS {
        assert_eq!(
            (list.held)(&read_back),
            (list.held)(&config),
            "`[{}]` came back as something other than what was written",
            list.section
        );
    }
}

/// The built-in lists, as the app starts with them, are the lists a configuration that has just
/// been made holds — which is what the lists reset promises, and what it used to reach by way of
/// a default configuration.
#[test]
fn the_lists_reset_puts_the_built_in_lists_back() {
    let mut config = AppConfig::default();
    for list in LISTS {
        (list.set)(&mut config, vec!["rhn-test".to_string()]);
    }

    reset_built_in(&mut config);

    for list in LISTS {
        assert_eq!(
            (list.held)(&config),
            &list.built_in()[..],
            "`[{}]` was not put back as the built-in list",
            list.section
        );
    }
}

/// The settings reset has to leave the lists exactly as they are, which is the whole of what
/// taking them out and putting them back is for.
#[test]
fn the_settings_reset_leaves_the_lists_exactly_as_they_are() {
    let mut config = AppConfig::default();
    for list in LISTS {
        (list.set)(&mut config, vec!["rhn-test".to_string()]);
    }
    let edited = held(&config);

    let taken = held(&config);
    let is_first_run = config.is_first_run;
    config = AppConfig::default();
    config.is_first_run = is_first_run;
    put(&mut config, taken);

    assert_eq!(
        held(&config),
        edited,
        "a reset of the settings moved a list"
    );
}

/// A kind added to the app is a row added to the table, and a list no row names would be a
/// field nobody wrote: empty on a fresh configuration, with nothing anywhere to say so.
#[test]
fn every_row_names_a_list_a_fresh_configuration_holds() {
    let config = AppConfig::default();

    for list in LISTS {
        assert_eq!(
            (list.held)(&config),
            &list.built_in()[..],
            "`[{}]` starts as something other than its built-in list, so the row does not name \
             the field the configuration holds it in",
            list.section
        );
    }
}

/// The lists are keyed by the section they are written under, and two of them may share a
/// section only by being two keys of it — the text kind's extensions and its names, which is the
/// only pair anywhere that does.
#[test]
fn a_list_is_keyed_by_its_section_and_its_key() {
    let mut keys: Vec<(&str, &str)> = LISTS.iter().map(|list| (list.section, list.key)).collect();
    let count = keys.len();
    keys.sort_unstable();
    keys.dedup();
    assert_eq!(
        keys.len(),
        count,
        "two rows are written under one section and key"
    );

    let shared: Vec<&str> = LISTS
        .iter()
        .map(|list| list.section)
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .filter(|section| LISTS.iter().filter(|list| list.section == *section).count() > 1)
        .collect();
    assert_eq!(
        shared,
        vec!["text"],
        "one section holds two lists and it is the text one"
    );
}
