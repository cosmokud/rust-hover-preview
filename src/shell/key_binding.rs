//! What a key setting names: one key, or alternatives of keys held together.
//!
//! `config.ini` writes a binding as alternatives separated by `,` and the keys of one
//! alternative separated by `+`: `space,page_down`, `ctrl+alt+space`, `ctrl+mouse4`. The
//! last key of an alternative is its *main* key, the one whose press is the edge a pin
//! counts, while the keys before it only have to be held. A single key is the one-token
//! case of the same thing, which is what keeps every older spelling working.
//!
//! Mouse buttons are named `mouse3`, `mouse4` and `mouse5`, and are carried as their
//! Windows virtual keys so that one down-state read answers for both kinds of key. The
//! buttons a window needs for itself — left and right — are not names this module takes.

use windows::Win32::UI::Input::KeyboardAndMouse::GetAsyncKeyState;

use super::key_input::key_to_vk;

/// The pin key an empty or unreadable setting falls back to.
pub(crate) const DEFAULT_PIN_KEY: &str = "space";

/// And the trigger key's.
pub(crate) const DEFAULT_TRIGGER_KEY: &str = "alt";

/// How many alternatives one binding may name, which is also how many chords the pin
/// hook keeps a down state for.
pub(crate) const MAX_ALTERNATIVES: usize = 8;

/// How many keys one alternative may name, main key included.
pub(crate) const MAX_CHORD_KEYS: usize = 6;

/// One alternative: the keys in the order they were written, the main key last.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Alternative {
    keys: Vec<i32>,
}

impl Alternative {
    /// The key whose press is the edge this alternative is counted on.
    fn main(&self) -> i32 {
        *self.keys.last().expect("a parsed alternative has a key")
    }

    /// The keys that only have to be held for it.
    fn companions(&self) -> &[i32] {
        &self.keys[..self.keys.len() - 1]
    }
}

/// A whole setting: one or more alternatives, any of which is the binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Binding {
    alternatives: Vec<Alternative>,
}

/// The virtual key one token names, with `mouse3` through `mouse5` taken before the
/// keyboard's own spellings.
pub(crate) fn parse_key(token: &str) -> Option<i32> {
    let token = token.trim().to_ascii_lowercase();
    match token.as_str() {
        "mouse3" => Some(0x04), // VK_MBUTTON
        "mouse4" => Some(0x05), // VK_XBUTTON1
        "mouse5" => Some(0x06), // VK_XBUTTON2
        _ => key_to_vk(&token),
    }
}

/// Parse a binding, or `None` if any alternative or token is not a key.
pub(crate) fn parse_binding(s: &str) -> Option<Binding> {
    let mut alternatives = Vec::new();
    for alternative in s.split(',') {
        let keys = alternative
            .split('+')
            .map(parse_key)
            .collect::<Option<Vec<_>>>()?;
        if keys.is_empty() || keys.len() > MAX_CHORD_KEYS {
            return None;
        }
        alternatives.push(Alternative { keys });
    }
    if alternatives.len() > MAX_ALTERNATIVES {
        return None;
    }
    Some(Binding { alternatives })
}

/// Parse a setting, falling back to a built-in default when it is empty or unreadable.
pub(crate) fn parse_or_default(s: &str, default: &str) -> Binding {
    parse_binding(s)
        .or_else(|| parse_binding(default))
        .expect("the built-in default is a key name")
}

/// Whether every key of any alternative is held, asked of `is_down`.
pub(crate) fn binding_is_active_with(b: &Binding, is_down: impl Fn(i32) -> bool) -> bool {
    b.alternatives
        .iter()
        .any(|alternative| alternative.keys.iter().all(|&vk| is_down(vk)))
}

/// Whether every key of any alternative is currently held.
pub(crate) fn binding_is_active(b: &Binding) -> bool {
    binding_is_active_with(b, key_is_down)
}

/// Whether `vk` is the main key of an alternative whose companions are all held, asked
/// of `is_down`. A main key also listed earlier is still the last token's key, and being
/// held itself is answered by the caller.
#[cfg(test)]
pub(crate) fn press_matches_with(b: &Binding, vk: i32, is_down: impl Fn(i32) -> bool) -> bool {
    b.alternatives.iter().any(|alternative| {
        alternative.main() == vk && alternative.companions().iter().all(|&key| is_down(key))
    })
}

/// Whether a keyboard event for `vk` is the press edge of an alternative.
#[cfg(test)]
pub(crate) fn binding_press_matches(b: &Binding, vk: i32) -> bool {
    press_matches_with(b, vk, key_is_down)
}

/// Whether a virtual key is one of the mouse buttons a binding may name.
pub(crate) fn is_mouse_key(vk: i32) -> bool {
    matches!(vk, 0x04..=0x06)
}

/// One alternative packed into a number the hook can read without a lock: main key in the
/// low byte, companion count above it, then one byte per companion.
pub(crate) fn packed_chord(alternative: &Alternative) -> u64 {
    let mut packed = alternative.main() as u64 & 0xFF;
    packed |= (alternative.companions().len() as u64) << 8;
    for (index, &vk) in alternative.companions().iter().enumerate() {
        packed |= (vk as u64 & 0xFF) << (12 + index * 8);
    }
    packed
}

/// A whole binding packed into the fixed table the pin hook reads, zero for unused slots.
pub(crate) fn packed_binding(b: &Binding) -> [u64; MAX_ALTERNATIVES] {
    let mut slots = [0; MAX_ALTERNATIVES];
    for (slot, alternative) in slots.iter_mut().zip(&b.alternatives) {
        *slot = packed_chord(alternative);
    }
    slots
}

/// The main key of a packed chord, or zero for an unused slot.
pub(crate) fn packed_main(packed: u64) -> i32 {
    (packed & 0xFF) as i32
}

/// Whether a packed chord's main key is `vk` and its companions are all held.
pub(crate) fn packed_press_matches(packed: u64, vk: i32) -> bool {
    packed != 0 && packed_main(packed) == vk && packed_companions_held(packed)
}

/// Whether every companion of a packed chord is currently held.
pub(crate) fn packed_companions_held(packed: u64) -> bool {
    let count = ((packed >> 8) & 0xF) as usize;
    (0..count).all(|index| key_is_down(((packed >> (12 + index * 8)) & 0xFF) as i32))
}

pub(crate) fn key_is_down(vk: i32) -> bool {
    unsafe { GetAsyncKeyState(vk) as u16 & 0x8000 != 0 }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn held(keys: &[i32]) -> impl Fn(i32) -> bool + '_ {
        move |vk| keys.contains(&vk)
    }

    #[test]
    fn a_single_key_is_a_binding_of_one_alternative() {
        let binding = parse_binding("space").expect("space is a key");
        assert_eq!(binding.alternatives.len(), 1);
        assert_eq!(binding.alternatives[0].main(), 0x20);
        assert!(binding.alternatives[0].companions().is_empty());

        assert!(press_matches_with(&binding, 0x20, held(&[0x20])));
        assert!(!binding_press_matches(&binding, 0x07));
        assert!(!press_matches_with(&binding, 0x11, held(&[0x11])));
    }

    #[test]
    fn a_comma_is_another_way_the_binding_can_be_matched() {
        let binding = parse_binding("space,page_down").expect("two keys");
        assert_eq!(binding.alternatives.len(), 2);
        assert!(press_matches_with(&binding, 0x20, held(&[0x20])));
        assert!(press_matches_with(&binding, 0x22, held(&[0x22])));
        assert!(!press_matches_with(&binding, 0x21, held(&[0x21])));
    }

    #[test]
    fn a_chord_is_main_last_and_companions_held() {
        let binding = parse_binding("ctrl+space").expect("a chord");
        assert_eq!(binding.alternatives[0].main(), 0x20);
        assert_eq!(binding.alternatives[0].companions(), &[0x11]);
        assert!(press_matches_with(&binding, 0x20, held(&[0x11, 0x20])));
        assert!(!press_matches_with(&binding, 0x20, held(&[0x20]),));
        assert!(!press_matches_with(&binding, 0x11, held(&[0x11, 0x20]),));

        let binding = parse_binding("ctrl+alt+space").expect("a longer chord");
        assert_eq!(binding.alternatives[0].main(), 0x20);
        assert!(press_matches_with(
            &binding,
            0x20,
            held(&[0x11, 0x12, 0x20])
        ));
        assert!(!press_matches_with(&binding, 0x20, held(&[0x11, 0x20])));
    }

    #[test]
    fn the_last_token_is_the_main_key_even_when_it_is_named_before() {
        let binding = parse_binding("ctrl+ctrl").expect("a chord");
        assert_eq!(binding.alternatives[0].main(), 0x11);
        assert_eq!(binding.alternatives[0].companions(), &[0x11]);
        assert!(press_matches_with(&binding, 0x11, held(&[0x11])));
    }

    #[test]
    fn mouse_buttons_three_to_five_are_keys_and_the_others_are_not() {
        let binding = parse_binding("space,mouse3").expect("a mouse key");
        assert!(press_matches_with(&binding, 0x20, held(&[0x20])));
        assert!(press_matches_with(&binding, 0x04, held(&[0x04])));
        assert!(!press_matches_with(&binding, 0x05, held(&[0x05])));

        let binding = parse_binding("ctrl+mouse4").expect("a mouse chord");
        assert!(press_matches_with(&binding, 0x05, held(&[0x11, 0x05])));
        assert!(!press_matches_with(&binding, 0x05, held(&[0x05])));
        assert_eq!(parse_key("mouse5"), Some(0x06));

        for rejected in ["mouse1", "mouse2", "middle", "bogus", ""] {
            assert_eq!(parse_binding(rejected), None, "{rejected:?} is not a key");
        }
    }

    #[test]
    fn whitespace_case_and_underscores_are_all_the_same_binding() {
        assert_eq!(
            parse_binding(" CTRL + Alt + SPACE "),
            parse_binding("ctrl+alt+space")
        );
        assert_eq!(parse_key("page_down"), Some(0x22));
        assert_eq!(parse_key("page_up"), Some(0x21));
        assert_eq!(parse_key("caps_lock"), Some(0x14));
        assert_eq!(parse_key("  Page_Down  "), Some(0x22));
    }

    #[test]
    fn an_empty_or_unknown_binding_falls_back_to_its_default() {
        assert_eq!(
            parse_or_default("", DEFAULT_PIN_KEY),
            parse_binding(DEFAULT_PIN_KEY).unwrap()
        );
        assert_eq!(
            parse_or_default("ctrl+nonsense", DEFAULT_TRIGGER_KEY),
            parse_binding(DEFAULT_TRIGGER_KEY).unwrap()
        );
        assert_eq!(
            parse_or_default("alt", DEFAULT_PIN_KEY),
            parse_binding("alt").unwrap()
        );
    }

    #[test]
    fn an_alternative_is_active_only_when_every_key_of_it_is_held() {
        let binding = parse_binding("ctrl+space,page_down").expect("two alternatives");
        assert!(binding_is_active_with(&binding, held(&[0x11, 0x20])));
        assert!(binding_is_active_with(&binding, held(&[0x22])));
        assert!(!binding_is_active_with(&binding, held(&[0x11])));
        assert!(!binding_is_active_with(&binding, held(&[0x20])));
        assert!(!binding_is_active_with(&binding, held(&[])));
    }

    #[test]
    fn packed_chords_read_back_as_the_alternative_they_came_from() {
        let binding = parse_binding("ctrl+alt+mouse4,space").expect("a binding");
        let packed = packed_binding(&binding);

        assert_eq!(packed_main(packed[0]), 0x05);
        assert_eq!(packed_main(packed[1]), 0x20);
        assert_eq!(packed_main(packed[2]), 0);
        assert!(!packed_press_matches(0, 0x20));
    }
}
