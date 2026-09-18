//! Playing an SVG: which parts of a document move, and what they are at a moment.
//!
//! usvg drops animation entirely — its own documentation says "no events and no
//! animations" — so a document that moves is played here as a sequence of still ones:
//! the declarations it carries (`<animate>`, `<animateTransform>`, `<set>`) are read
//! from its XML, the value each of them has at a moment is worked out, and the document
//! is written out again with those values in it for the renderer to draw. What that
//! costs is a parse and a rasterization per frame, which is what animation costs in a
//! renderer that has none.
//!
//! What is supported is the part of SMIL a document in a folder is likely to use:
//! numbers and colors interpolated between the values a declaration names, transforms
//! of the five function types, `keyTimes`, `dur`, `repeatCount` and `fill="freeze"`.
//! What is not is the rest — event-based timing (`begin="click"`), `additive` and
//! `accumulate`, `calcMode="spline"`, and interpolation of paths and lists — and a value
//! that cannot be blended is held until the next one instead of being refused, which is
//! what `calcMode="discrete"` means anyway. A declaration that asks for something
//! outside all of this is one animation fewer, and a document whose animations are all
//! beyond it is simply a still document, which is what it was before any of this
//! existed.

use roxmltree::{Attribute, Document, Node, NodeId, NodeType};
use std::collections::HashMap;

/// How long one drawn frame is shown for. A document is not drawn faster than this:
/// the player will not show a frame sooner than it, so drawing one would only spend a
/// thread.
pub const FRAME_MS: u32 = 33;

/// The longest cycle that is played, in frames — a little over ten seconds.
///
/// A preview is a glance, and a cycle of minutes is not one: past this the first ten
/// seconds are played and looped, which is the part of such a document a hover was ever
/// going to show.
const MAX_FRAMES: u32 = 300;

/// The animations a document declares, and how long one pass of them is.
pub struct Playback {
    animations: Vec<Animation>,
    frames: u32,
}

impl Playback {
    /// The animations `document` declares, or nothing when it declares none this can
    /// play — which is also the answer for a document that does not move at all.
    pub fn parse(document: &Document) -> Option<Self> {
        let animations: Vec<Animation> = document
            .descendants()
            .filter(Node::is_element)
            .filter_map(parse_animation)
            .collect();

        if animations.is_empty() {
            return None;
        }

        let cycle_seconds = animations
            .iter()
            .map(Animation::active_end)
            .fold(0.0f32, f32::max)
            .max(FRAME_MS as f32 / 1000.0);
        let frames =
            ((cycle_seconds * 1000.0 / FRAME_MS as f32).ceil() as u32).clamp(1, MAX_FRAMES);

        Some(Self { animations, frames })
    }

    /// How many frames one pass of this document is.
    pub fn frames(&self) -> u32 {
        self.frames
    }

    /// The document as XML, with every animated value at the moment `frame` falls on.
    pub fn document_at(&self, document: &Document, frame: u32) -> String {
        let time = frame * FRAME_MS;

        let mut overrides: HashMap<NodeId, Resolved> = HashMap::new();
        for animation in &self.animations {
            let Some(value) = animation.value_at(time) else {
                continue;
            };

            let resolved = overrides.entry(animation.target).or_default();
            match &animation.kind {
                Kind::Transform(_) => resolved.transform = Some(value),
                Kind::Attribute(name) => resolved.attributes.push((name.clone(), value)),
            }
        }

        write_document(document, &overrides)
    }
}

/// One declaration: one attribute of one element, over time.
struct Animation {
    target: NodeId,
    kind: Kind,
    timing: Timing,
    key_times: Vec<f32>,
    values: Values,
}

/// What a declaration writes when it is applied.
enum Kind {
    /// The element's `transform`, written as the function the declaration names.
    Transform(Transform),
    /// An attribute of the element, by name.
    Attribute(String),
}

/// The transform functions `animateTransform` can animate.
#[derive(Clone, Copy)]
enum Transform {
    Rotate,
    Scale,
    Translate,
    SkewX,
    SkewY,
}

impl Transform {
    fn parse(name: &str) -> Option<Self> {
        match name.trim() {
            "rotate" => Some(Self::Rotate),
            "scale" => Some(Self::Scale),
            "translate" => Some(Self::Translate),
            "skewX" => Some(Self::SkewX),
            "skewY" => Some(Self::SkewY),
            _ => None,
        }
    }

    /// The function as SVG spells it, with the numbers the declaration named.
    fn write(self, numbers: &[f32]) -> String {
        let name = match self {
            Self::Rotate => "rotate",
            Self::Scale => "scale",
            Self::Translate => "translate",
            Self::SkewX => "skewX",
            Self::SkewY => "skewY",
        };

        let mut text = String::from(name);
        text.push('(');
        for (index, number) in numbers.iter().enumerate() {
            if index > 0 {
                text.push(' ');
            }
            text.push_str(&format_number(*number));
        }
        text.push(')');

        text
    }
}

/// When a declaration runs, and for how long.
struct Timing {
    begin_seconds: f32,
    duration_seconds: f32,
    /// How many times it repeats, or `None` for an indefinite repeat.
    repeats: Option<u32>,
    /// Whether it holds its last value once it has run out.
    freeze: bool,
}

/// The values a declaration names, one per key time.
enum Values {
    Numbers(Vec<Vec<f32>>),
    Colors(Vec<[f32; 3]>),
    Steps(Vec<String>),
}

/// A value at a moment, in the shape the declaration's target takes.
enum Value {
    Numbers(Vec<f32>),
    Text(String),
}

/// What is written over an element: its transform, and the attributes a declaration
/// targets by name.
#[derive(Default)]
struct Resolved {
    transform: Option<String>,
    attributes: Vec<(String, String)>,
}

impl Animation {
    /// The value this declaration has at `time` milliseconds into the document, or
    /// nothing when it does not apply then — before it begins, or after it has run out
    /// without being frozen — in which case the attribute keeps the value the document
    /// gave it.
    fn value_at(&self, time_ms: u32) -> Option<String> {
        let time = time_ms as f32 / 1000.0;
        let local = time - self.timing.begin_seconds;
        if local < 0.0 {
            return None;
        }

        let duration = self.timing.duration_seconds.max(0.001);
        let active = match self.timing.repeats {
            Some(repeats) => duration * repeats as f32,
            None => f32::INFINITY,
        };

        let phase = if local >= active {
            if !self.timing.freeze {
                return None;
            }
            1.0
        } else {
            (local % duration) / duration
        };

        let value = self.values.at(&self.key_times, phase)?;

        Some(match (&self.kind, value) {
            (Kind::Transform(transform), Value::Numbers(numbers)) => transform.write(&numbers),
            (Kind::Attribute(_), Value::Numbers(numbers)) => numbers
                .iter()
                .map(|number| format_number(*number))
                .collect::<Vec<String>>()
                .join(" "),
            (_, Value::Text(text)) => text,
        })
    }

    /// When this declaration has run its course, repeats included.
    fn active_end(&self) -> f32 {
        let run = match self.timing.repeats {
            Some(repeats) => self.timing.duration_seconds * repeats as f32,
            None => self.timing.duration_seconds,
        };

        self.timing.begin_seconds + run
    }
}

impl Values {
    /// The value at `phase`, which runs 0 to 1 across one pass. A value that cannot be
    /// blended with the one before it is held rather than guessed at.
    fn at(&self, key_times: &[f32], phase: f32) -> Option<Value> {
        match self {
            Self::Numbers(values) => {
                let (first, second, blend) = segment(key_times, phase, values.len());
                let (Some(from), Some(to)) = (values.get(first), values.get(second)) else {
                    return None;
                };

                if from.len() != to.len() {
                    return Some(Value::Numbers(from.clone()));
                }

                Some(Value::Numbers(
                    from.iter()
                        .zip(to.iter())
                        .map(|(from, to)| from + ((to - from) * blend))
                        .collect(),
                ))
            }
            Self::Colors(values) => {
                let (first, second, blend) = segment(key_times, phase, values.len());
                let (Some(from), Some(to)) = (values.get(first), values.get(second)) else {
                    return None;
                };

                let mixed = [0, 1, 2].map(|channel| {
                    let value = from[channel] + ((to[channel] - from[channel]) * blend);
                    (value * 255.0).round().clamp(0.0, 255.0) as u8
                });

                Some(Value::Text(format!(
                    "#{:02x}{:02x}{:02x}",
                    mixed[0], mixed[1], mixed[2]
                )))
            }
            Self::Steps(values) => {
                let (first, _, _) = segment(key_times, phase, values.len());
                values.get(first).cloned().map(Value::Text)
            }
        }
    }
}

/// The pair of values `phase` falls between, and how far between them it is.
fn segment(key_times: &[f32], phase: f32, count: usize) -> (usize, usize, f32) {
    if count == 0 {
        return (0, 0, 0.0);
    }
    if count == 1 || key_times.len() < 2 {
        return (0, 0, 0.0);
    }

    for index in 0..key_times.len() - 1 {
        let (from, to) = (key_times[index], key_times[index + 1]);
        if phase <= to || index == key_times.len() - 2 {
            let span = (to - from).max(f32::EPSILON);

            return (index, index + 1, ((phase - from) / span).clamp(0.0, 1.0));
        }
    }

    (key_times.len() - 1, key_times.len() - 1, 0.0)
}

/// One animation declaration, or nothing when it asks for something this cannot play.
fn parse_animation(node: Node) -> Option<Animation> {
    let target = node.parent_element()?.id();
    let tag = node.tag_name().name();

    let timing = Timing {
        begin_seconds: node.attribute("begin").and_then(parse_begin).unwrap_or(0.0),
        duration_seconds: node
            .attribute("dur")
            .and_then(parse_seconds)
            .unwrap_or(1.0)
            .max(0.001),
        repeats: parse_repeats(node.attribute("repeatCount")),
        freeze: node.attribute("fill") == Some("freeze"),
    };

    let kind = match tag {
        "animateTransform" => {
            Kind::Transform(Transform::parse(node.attribute("type").unwrap_or(""))?)
        }
        // A `set` holds a value rather than moving between them, which is one value and
        // a step between it and itself.
        _ => Kind::Attribute(node.attribute("attributeName")?.trim().to_string()),
    };

    let named: Vec<String> = if tag == "set" {
        vec![node.attribute("to")?.to_string()]
    } else {
        named_values(node)?
    };

    let key_times = parse_key_times(node.attribute("keyTimes"), named.len());
    let values = match &kind {
        Kind::Transform(_) => Values::Numbers(
            named
                .iter()
                .map(|value| parse_numbers(value))
                .collect::<Option<Vec<_>>>()?,
        ),
        Kind::Attribute(name) => values_for_attribute(name, &named)?,
    };

    Some(Animation {
        target,
        kind,
        timing,
        key_times,
        values,
    })
}

/// The values an `animate` names: its `values` list, or the pair `from` and `to`.
fn named_values(node: Node) -> Option<Vec<String>> {
    if let Some(values) = node.attribute("values") {
        let values: Vec<String> = values
            .split(';')
            .map(|value| value.trim().to_string())
            .filter(|value| !value.is_empty())
            .collect();

        return (values.len() > 1).then_some(values);
    }

    let from = node.attribute("from")?.trim();
    let to = node.attribute("to")?.trim();

    Some(vec![from.to_string(), to.to_string()])
}

/// Which of the three value shapes an attribute of this name takes.
fn values_for_attribute(name: &str, named: &[String]) -> Option<Values> {
    let attribute = name.to_ascii_lowercase();

    if COLOR_ATTRIBUTES.contains(&attribute.as_str()) {
        if let Some(colors) = named
            .iter()
            .map(|value| parse_color(value))
            .collect::<Option<Vec<_>>>()
        {
            return Some(Values::Colors(colors));
        }
    }

    if NUMERIC_ATTRIBUTES.contains(&attribute.as_str()) {
        if let Some(numbers) = named
            .iter()
            .map(|value| parse_numbers(value))
            .collect::<Option<Vec<_>>>()
        {
            return Some(Values::Numbers(numbers));
        }
    }

    Some(Values::Steps(
        named.iter().map(|value| value.trim().to_string()).collect(),
    ))
}

/// The attributes a hover's documents animate as numbers, and as colors.
const NUMERIC_ATTRIBUTES: &[&str] = &[
    "cx",
    "cy",
    "fill-opacity",
    "font-size",
    "height",
    "offset",
    "opacity",
    "r",
    "rx",
    "ry",
    "stroke-dashoffset",
    "stroke-miterlimit",
    "stroke-opacity",
    "stroke-width",
    "width",
    "x",
    "y",
];
const COLOR_ATTRIBUTES: &[&str] = &["color", "fill", "flood-color", "stop-color", "stroke"];

/// Where each value sits in one pass, 0 to 1: the `keyTimes` a declaration gives, or
/// values spread evenly across the pass where it gives none.
fn parse_key_times(key_times: Option<&str>, count: usize) -> Vec<f32> {
    if let Some(key_times) = key_times {
        let parsed: Vec<f32> = key_times
            .split(';')
            .filter_map(|time| time.trim().parse::<f32>().ok())
            .collect();

        if parsed.len() == count {
            return parsed;
        }
    }

    if count <= 1 {
        return vec![0.0];
    }

    (0..count)
        .map(|index| index as f32 / (count - 1) as f32)
        .collect()
}

/// When a declaration begins: the first moment in its `begin` list that is a time.
/// An event — `click`, `mouseover` — is not something a hover can wait for, so a list
/// with nothing but events in it begins at once rather than never.
fn parse_begin(begin: &str) -> Option<f32> {
    begin
        .split(';')
        .filter_map(|entry| parse_seconds(entry.trim()))
        .next()
}

/// A clock value: `2s`, `500ms`, or a bare number of seconds.
fn parse_seconds(text: &str) -> Option<f32> {
    if let Some(milliseconds) = text.strip_suffix("ms") {
        return milliseconds.trim().parse().ok();
    }

    if let Some(seconds) = text.strip_suffix('s') {
        return seconds.trim().parse().ok();
    }

    text.parse().ok()
}

/// How many times a declaration repeats, or `None` for an indefinite repeat.
fn parse_repeats(repeat_count: Option<&str>) -> Option<u32> {
    let repeat_count = repeat_count?.trim();

    if repeat_count.eq_ignore_ascii_case("indefinite") {
        return None;
    }

    let repeats = repeat_count.parse::<f32>().ok()?.max(1.0);

    Some(repeats.ceil() as u32)
}

fn parse_numbers(value: &str) -> Option<Vec<f32>> {
    let numbers: Vec<f32> = value
        .split([',', ' ', '\t', '\n'])
        .filter(|part| !part.trim().is_empty())
        .map(|part| part.trim().parse::<f32>().ok())
        .collect::<Option<Vec<f32>>>()?;

    (!numbers.is_empty()).then_some(numbers)
}

/// A color as three channels, from the spellings a document is likely to use.
fn parse_color(value: &str) -> Option<[f32; 3]> {
    let value = value.trim();

    if let Some(hex) = value.strip_prefix('#') {
        let digits: Vec<u8> = hex
            .chars()
            .map(|digit| digit.to_digit(16).map(|digit| digit as u8))
            .collect::<Option<Vec<u8>>>()?;

        return match digits.len() {
            3 => Some([
                digits[0] as f32 / 15.0,
                digits[1] as f32 / 15.0,
                digits[2] as f32 / 15.0,
            ]),
            6 => Some([
                (digits[0] * 16 + digits[1]) as f32 / 255.0,
                (digits[2] * 16 + digits[3]) as f32 / 255.0,
                (digits[4] * 16 + digits[5]) as f32 / 255.0,
            ]),
            _ => None,
        };
    }

    let lower = value.to_ascii_lowercase();
    let inside = lower
        .strip_prefix("rgb(")
        .or_else(|| lower.strip_prefix("rgba("))?
        .strip_suffix(')')?;

    let channels: Vec<f32> = inside
        .split([',', ' ', '/'])
        .filter(|part| !part.trim().is_empty())
        .take(3)
        .map(|part| part.trim().parse::<f32>())
        .collect::<Result<Vec<f32>, _>>()
        .ok()?;

    (channels.len() == 3).then(|| {
        [
            (channels[0] / 255.0).clamp(0.0, 1.0),
            (channels[1] / 255.0).clamp(0.0, 1.0),
            (channels[2] / 255.0).clamp(0.0, 1.0),
        ]
    })
}

/// A number as SVG writes it: no trailing zeros, and no exponent.
fn format_number(value: f32) -> String {
    let mut text = format!("{value:.3}");

    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    if text.is_empty() || text == "-0" {
        text.push('0');
    }

    text
}

/// The document as XML, with the values that are animated at this moment written into
/// it. Everything else is written as it was found, which is what keeps a document that
/// is mostly not animated exactly the one that would have been drawn anyway.
fn write_document(document: &Document, overrides: &HashMap<NodeId, Resolved>) -> String {
    let mut out = String::new();

    if let Some(root) = document.root().first_element_child() {
        write_element(root, overrides, &mut out);
    }

    out
}

fn write_element(node: Node, overrides: &HashMap<NodeId, Resolved>, out: &mut String) {
    let resolved = overrides.get(&node.id());
    let replaced = |name: &str| {
        resolved.is_some_and(|resolved| {
            resolved.attributes.iter().any(|(named, _)| named == name)
                || (name == "transform" && resolved.transform.is_some())
        })
    };

    out.push('<');
    out.push_str(&element_name(node));

    // A namespace is a declaration rather than an attribute to the reader, and it is
    // written back as one: a document that loses its `xmlns` is not the document that
    // was parsed.
    for namespace in node.namespaces() {
        out.push_str(&match namespace.name() {
            Some(prefix) => format!(" xmlns:{prefix}=\""),
            None => " xmlns=\"".to_string(),
        });
        write_escaped(namespace.uri(), out);
        out.push('"');
    }

    for attribute in node.attributes() {
        let name = attribute_name(node, attribute);
        if replaced(&name) {
            continue;
        }

        out.push(' ');
        out.push_str(&name);
        out.push_str("=\"");
        write_escaped(attribute.value(), out);
        out.push('"');
    }

    if let Some(resolved) = resolved {
        if let Some(transform) = &resolved.transform {
            // The element's own transform comes first, so the animated one composes
            // over it the way a second transform on the same element would.
            out.push_str(" transform=\"");
            if let Some(own) = node.attribute("transform") {
                write_escaped(own, out);
                out.push(' ');
            }
            write_escaped(transform, out);
            out.push('"');
        }

        for (name, value) in &resolved.attributes {
            out.push(' ');
            out.push_str(name);
            out.push_str("=\"");
            write_escaped(value, out);
            out.push('"');
        }
    }

    if node.children().next().is_none() {
        out.push_str("/>");
        return;
    }

    out.push('>');

    for child in node.children() {
        match child.node_type() {
            NodeType::Element => write_element(child, overrides, out),
            NodeType::Text => {
                if let Some(text) = child.text() {
                    write_escaped(text, out);
                }
            }
            // Comments and processing instructions draw nothing and are dropped: what
            // is written is a document for the renderer.
            _ => {}
        }
    }

    out.push_str("</");
    out.push_str(&element_name(node));
    out.push('>');
}

/// The element's name as it was written, prefix and all.
fn element_name(node: Node) -> String {
    let name = node.tag_name().name();

    match node
        .tag_name()
        .namespace()
        .and_then(|namespace| node.lookup_prefix(namespace))
    {
        Some(prefix) => format!("{prefix}:{name}"),
        None => name.to_string(),
    }
}

/// The attribute's name as it was written, prefix and all — including a namespace
/// declaration, which is what keeps a document's own prefixes intact.
fn attribute_name(node: Node, attribute: Attribute) -> String {
    if attribute.namespace() == Some("http://www.w3.org/2000/xmlns/") {
        return if attribute.name() == "xmlns" {
            "xmlns".to_string()
        } else {
            format!("xmlns:{}", attribute.name())
        };
    }

    match attribute
        .namespace()
        .and_then(|namespace| node.lookup_prefix(namespace))
    {
        Some(prefix) => format!("{prefix}:{}", attribute.name()),
        None => attribute.name().to_string(),
    }
}

fn write_escaped(text: &str, out: &mut String) {
    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            other => out.push(other),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The document and its playback, which is the pair every test works with.
    fn playback(xml: &str) -> (Document<'_>, Playback) {
        let document = Document::parse(xml).expect("a parsed document");
        let playback = Playback::parse(&document).expect("a document that moves");

        (document, playback)
    }

    #[test]
    fn a_document_that_does_not_move_has_nothing_to_play() {
        let document = Document::parse(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>"#,
        )
        .expect("a parsed document");

        assert!(Playback::parse(&document).is_none());
    }

    /// A transform is written into the element's own `transform`, over whatever it
    /// already had, and the rest of the document comes through untouched.
    #[test]
    fn writes_a_transform_at_the_moment_it_is_asked_for() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><g transform="translate(5 5)"><rect width="10" height="10"><animateTransform attributeName="transform" type="rotate" from="0" to="360" dur="1s" repeatCount="indefinite"/></rect></g></svg>"#,
        );

        let first = playback.document_at(&document, 0);
        assert!(first.contains(r#"transform="rotate(0)""#), "{first}");
        assert!(first.contains(r#"transform="translate(5 5)""#), "{first}");

        // Frame fifteen of a one-second pass is 495 milliseconds in, which is 178.2
        // degrees of the turn.
        let middle = playback.document_at(&document, 15);
        assert!(middle.contains(r#"transform="rotate(178.2)""#), "{middle}");
    }

    #[test]
    fn interpolates_an_opacity_between_its_values() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10" opacity="1"><animate attributeName="opacity" values="0;1;0" keyTimes="0;0.5;1" dur="2s" repeatCount="indefinite"/></rect></svg>"#,
        );

        let first = playback.document_at(&document, 0);
        assert!(first.contains(r#"opacity="0""#), "{first}");

        // 495 milliseconds into a two-second pass is 495 thousandths of the way to the
        // middle value.
        let quarter = playback.document_at(&document, 15);
        assert!(quarter.contains(r#"opacity="0.495""#), "{quarter}");
    }

    #[test]
    fn blends_a_color_between_its_values() {
        let (document, playback) = playback(
            r##"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10" fill="#000000"><animate attributeName="fill" from="#000000" to="#ffffff" dur="1s"/></rect></svg>"##,
        );

        let first = playback.document_at(&document, 0);
        assert!(first.contains(r##"fill="#000000""##), "{first}");

        let middle = playback.document_at(&document, 15);
        assert!(middle.contains(r##"fill="#7e7e7e""##), "{middle}");
    }

    /// A value that cannot be blended is held until the next one, which is what a
    /// `set` is by nature.
    #[test]
    fn holds_a_value_that_cannot_be_blended() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"><set attributeName="stroke-dasharray" to="4 2" begin="0s" dur="1s"/></rect></svg>"#,
        );

        assert!(playback
            .document_at(&document, 0)
            .contains(r#"stroke-dasharray="4 2""#));
    }

    /// Until a declaration begins, the element keeps the value the document gave it.
    #[test]
    fn leaves_an_element_alone_before_its_animation_begins() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10" opacity="1"><animate attributeName="opacity" from="0" to="1" begin="1s" dur="1s" repeatCount="indefinite"/></rect></svg>"#,
        );

        let before = playback.document_at(&document, 15);
        assert!(before.contains(r#"opacity="1""#), "{before}");

        // Frame thirty-one is 1.023 seconds in: twenty-three milliseconds into the
        // declaration's own pass.
        let after = playback.document_at(&document, 31);
        assert!(after.contains(r#"opacity="0.023""#), "{after}");
    }

    #[test]
    fn a_pass_is_never_longer_than_the_cap() {
        let (_document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"><animate attributeName="opacity" from="0" to="1" dur="600s" repeatCount="indefinite"/></rect></svg>"#,
        );

        assert_eq!(playback.frames(), MAX_FRAMES);
    }
}
