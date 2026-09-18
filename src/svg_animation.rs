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

/// Whether a document says it moves at all.
///
/// This is a question about the document rather than about this module: what a
/// document declares is reason enough to hand it to an engine that plays the whole of
/// SMIL and CSS, and this reader's own subset is what draws it when there is no such
/// engine on the machine. A declaration outside the subset is one the engine can still
/// play, which is why the check is for any declaration at all rather than for one this
/// can carry out.
pub fn declares_animation(document: &Document) -> bool {
    document.descendants().filter(Node::is_element).any(|node| {
        matches!(
            node.tag_name().name(),
            "animate" | "animateTransform" | "animateMotion" | "set"
        ) || (node.tag_name().name() == "style"
            && node.text().is_some_and(|text| text.contains("@keyframes")))
    })
}

/// The animations a document declares, and how long one pass of them is.
pub struct Playback {
    animations: Vec<Animation>,
    frames: u32,
    /// Whether the document's own pass is longer than the frames a preview plays of
    /// it: a clock's hour hand takes twelve hours to come round, and what a hover shows
    /// of that is the first stretch of it rather than the whole turn.
    truncated: bool,
}

impl Playback {
    /// The animations `document` declares, or nothing when it declares none this can
    /// play — which is also the answer for a document that does not move at all.
    ///
    /// Two front ends feed the one model: the declarations in the markup, and the
    /// stylesheets that name `@keyframes`. Where both reach the same attribute of the
    /// same element the stylesheet is written last, so it is the one that shows, which
    /// is the order a browser's cascade puts them in.
    pub fn parse(document: &Document) -> Option<Self> {
        let mut animations: Vec<Animation> = document
            .descendants()
            .filter(Node::is_element)
            .filter_map(parse_animation)
            .collect();
        animations.extend(stylesheet_animations(document));

        if animations.is_empty() {
            return None;
        }

        let cycle_seconds = animations
            .iter()
            .map(Animation::active_end)
            .fold(0.0f32, f32::max)
            .max(FRAME_MS as f32 / 1000.0);
        let pass_frames = (cycle_seconds * 1000.0 / FRAME_MS as f32).ceil() as u32;
        let frames = pass_frames.clamp(1, MAX_FRAMES);

        Some(Self {
            animations,
            frames,
            truncated: pass_frames > MAX_FRAMES,
        })
    }

    /// How many frames one pass of this document is.
    pub fn frames(&self) -> u32 {
        self.frames
    }

    /// Whether the pass that is played is the whole of what the document says. A pass
    /// that was cut short is played forward rather than looped — a clock's second hand
    /// keeps going instead of jumping back to where the hover found it — while one that
    /// holds everything the document does is repeated, as an animation of two seconds
    /// is meant to be.
    pub fn repeats_its_pass(&self) -> bool {
        !self.truncated
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
                Kind::Transform(_) | Kind::TransformList { .. } => {
                    resolved.transform = Some(value);
                }
                Kind::Attribute(name) => resolved.set_attribute(name, value),
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
    /// The element's `transform`, as a stylesheet writes it: a list of functions that
    /// are the same in every value, with their arguments blended between them. The
    /// origin is what CSS rotates and scales about, written into the transform as the
    /// pair of translations that move it there and back.
    TransformList {
        functions: Vec<TransformFunction>,
        origin: Option<(f32, f32)>,
    },
    /// An attribute of the element, by name.
    Attribute(String),
}

/// One CSS transform function and how many numbers it takes.
#[derive(Clone, Copy)]
struct TransformFunction {
    name: CssTransform,
    arity: usize,
}

#[derive(Clone, Copy)]
enum CssTransform {
    Translate,
    Scale,
    Rotate,
    SkewX,
    SkewY,
    Matrix,
}

impl CssTransform {
    fn as_str(self) -> &'static str {
        match self {
            Self::Translate => "translate",
            Self::Scale => "scale",
            Self::Rotate => "rotate",
            Self::SkewX => "skewX",
            Self::SkewY => "skewY",
            Self::Matrix => "matrix",
        }
    }

    /// The numbers that make this function do nothing, for a keyframe that says
    /// `transform: none` between two that do not.
    fn identity(self) -> &'static [f32] {
        match self {
            Self::Translate => &[0.0, 0.0],
            Self::Scale => &[1.0, 1.0],
            Self::Rotate => &[0.0],
            Self::SkewX | Self::SkewY => &[0.0],
            Self::Matrix => &[1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
        }
    }
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

impl Resolved {
    /// Write an attribute over the element, replacing anything already written for it:
    /// two declarations that reach the same attribute are one attribute in the
    /// document, and the last one asked for is the one that shows.
    fn set_attribute(&mut self, name: &str, value: String) {
        if let Some(existing) = self
            .attributes
            .iter_mut()
            .find(|(written, _)| written == name)
        {
            existing.1 = value;
            return;
        }

        self.attributes.push((name.to_string(), value));
    }
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
            (Kind::TransformList { functions, origin }, Value::Numbers(numbers)) => {
                write_transform_list(functions, *origin, &numbers)
            }
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

/// The element a declaration animates: the one its `href` names, or the one it is
/// written inside.
///
/// A document is free to keep a declaration apart from what it moves — a clock puts
/// its three hands at the end of the file and the three animations that turn them in
/// `<defs>`, pointing at each hand by id — so a declaration that names an element is
/// about that element, and not about the `<defs>` it happens to sit in. Both spellings
/// of the attribute are read: `href` and the `xlink:href` every older document uses.
fn target_element<'a, 'input>(node: Node<'a, 'input>) -> Option<Node<'a, 'input>> {
    let href = node
        .attribute("href")
        .or_else(|| node.attribute(("http://www.w3.org/1999/xlink", "href")));

    if let Some(href) = href {
        if let Some(id) = href.trim().strip_prefix('#') {
            let named = node
                .document()
                .descendants()
                .find(|element| element.attribute("id") == Some(id));

            if named.is_some() {
                return named;
            }
        }
    }

    node.parent_element()
}

/// One animation declaration, or nothing when it asks for something this cannot play.
fn parse_animation(node: Node) -> Option<Animation> {
    let target = target_element(node)?.id();
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
        // A transform list is what a stylesheet writes, and is built by the stylesheet
        // reader rather than from a declaration in the markup.
        Kind::TransformList { .. } => return None,
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

/// The properties a keyframe block may animate, which are the ones this can carry into
/// a document as attributes.
const KEYFRAME_PROPERTIES: &[&str] = &[
    "fill",
    "fill-opacity",
    "opacity",
    "stroke",
    "stroke-opacity",
    "stroke-width",
    "transform",
];

/// The stylesheets of a document, read into the same model the declarations in its
/// markup are read into.
fn stylesheet_animations(document: &Document) -> Vec<Animation> {
    let viewport = viewport_size(document);
    let mut stylesheet = Stylesheet::default();

    for style in document
        .descendants()
        .filter(|node| node.has_tag_name("style"))
    {
        if let Some(text) = style.text() {
            stylesheet.absorb(text);
        }
    }

    if stylesheet.keyframes.is_empty() {
        return Vec::new();
    }

    let mut animations = Vec::new();

    // Every rule against every element it matches. A document a hover meets is a small
    // one, so the plain sweep is cheaper than the bookkeeping a match cache would need.
    for (selector, body) in &stylesheet.rules {
        for element in document.descendants().filter(Node::is_element) {
            if !selector_matches(selector, element) {
                continue;
            }

            animations.extend(element_animations(
                element,
                &declarations(body),
                &stylesheet,
                viewport,
            ));
        }
    }

    // And what an element says about itself in its own `style` attribute.
    for element in document.descendants().filter(Node::is_element) {
        if let Some(style) = element.attribute("style") {
            animations.extend(element_animations(
                element,
                &declarations(style),
                &stylesheet,
                viewport,
            ));
        }
    }

    animations
}

/// The styles of a document: the `@keyframes` blocks it names, and the rules that put
/// them on elements.
#[derive(Default)]
struct Stylesheet {
    keyframes: HashMap<String, Vec<(f32, String)>>,
    rules: Vec<(String, String)>,
}

impl Stylesheet {
    fn absorb(&mut self, css: &str) {
        for (prelude, body) in blocks(&strip_comments(css)) {
            let lower = prelude.to_ascii_lowercase();

            if let Some(name) = lower.strip_prefix("@keyframes") {
                let name = name.trim().to_string();
                let frames = keyframe_frames(&body);

                if !name.is_empty() && !frames.is_empty() {
                    self.keyframes.insert(name, frames);
                }
            } else if !prelude.starts_with('@') {
                self.rules.push((prelude, body));
            }
        }
    }
}

/// The steps of one `@keyframes` body, as the offset each sits at and what it declares
/// there.
fn keyframe_frames(body: &str) -> Vec<(f32, String)> {
    let mut frames: Vec<(f32, String)> = Vec::new();

    for (prelude, declarations) in blocks(body) {
        for offset in prelude.split(',') {
            let Some(offset) = keyframe_offset(offset.trim()) else {
                continue;
            };

            frames.push((offset, declarations.clone()));
        }
    }

    frames.sort_by(|(left, _), (right, _)| left.total_cmp(right));
    frames
}

/// The offset a keyframe selector names: `from`, `to`, or a percentage of the pass.
fn keyframe_offset(selector: &str) -> Option<f32> {
    match selector.to_ascii_lowercase().as_str() {
        "from" => Some(0.0),
        "to" => Some(1.0),
        _ => selector
            .strip_suffix('%')?
            .trim()
            .parse::<f32>()
            .ok()
            .map(|percent| (percent / 100.0).clamp(0.0, 1.0)),
    }
}

/// One animation a declaration block asks for.
struct AnimationRequest {
    name: String,
    duration: f32,
    delay: f32,
    repeats: Option<u32>,
    freeze: bool,
}

/// The animations a rule asks for, from the `animation` shorthand or from its
/// longhands.
fn animation_requests(declarations: &[(String, String)]) -> Vec<AnimationRequest> {
    let property = |name: &str| {
        declarations
            .iter()
            .find(|(property, _)| property == name)
            .map(|(_, value)| value.trim().to_string())
    };
    let comma_list = |value: String| -> Vec<String> {
        value
            .split(',')
            .map(|part| part.trim().to_string())
            .collect()
    };

    let mut requests = Vec::new();

    for value in declarations
        .iter()
        .filter(|(property, _)| property == "animation")
        .map(|(_, value)| value.clone())
    {
        for part in comma_list(value) {
            requests.push(shorthand_request(&part));
        }
    }

    if let Some(names) = property("animation-name") {
        let names = comma_list(names);
        let durations = property("animation-duration")
            .map(comma_list)
            .unwrap_or_default();
        let delays = property("animation-delay")
            .map(comma_list)
            .unwrap_or_default();
        let repeats = property("animation-iteration-count")
            .map(comma_list)
            .unwrap_or_default();
        let fills = property("animation-fill-mode")
            .map(comma_list)
            .unwrap_or_default();

        for (index, name) in names.iter().enumerate() {
            requests.push(AnimationRequest {
                name: name.clone(),
                duration: durations
                    .get(index)
                    .and_then(|value| parse_css_time(value))
                    .unwrap_or(0.0),
                delay: delays
                    .get(index)
                    .and_then(|value| parse_css_time(value))
                    .unwrap_or(0.0),
                repeats: repeats
                    .get(index)
                    .and_then(|value| parse_iterations(value))
                    .unwrap_or(Some(1)),
                freeze: fills
                    .get(index)
                    .is_some_and(|value| value.contains("forwards") || value.contains("both")),
            });
        }
    }

    requests.retain(|request| !request.name.is_empty() && request.duration > 0.0);
    requests
}

/// One animation out of the `animation` shorthand: `spin 1.2s linear infinite`, or
/// the same words in any order.
fn shorthand_request(part: &str) -> AnimationRequest {
    let mut request = AnimationRequest {
        name: String::new(),
        duration: 0.0,
        delay: 0.0,
        repeats: Some(1),
        freeze: false,
    };
    let mut times = 0;

    for token in part.split_whitespace() {
        let lower = token.to_ascii_lowercase();

        if let Some(seconds) = parse_css_time(token) {
            match times {
                0 => request.duration = seconds,
                1 => request.delay = seconds,
                _ => {}
            }
            times += 1;
            continue;
        }

        match lower.as_str() {
            "infinite" => request.repeats = None,
            "forwards" | "both" => request.freeze = true,
            // Timing functions, directions and play state are not played: an ease is
            // drawn linearly rather than not at all.
            "normal" | "reverse" | "alternate" | "alternate-reverse" | "none" | "backwards"
            | "running" | "paused" | "linear" | "ease" | "ease-in" | "ease-out" | "ease-in-out"
            | "step-start" | "step-end" => {}
            _ if lower.starts_with("steps(") || lower.starts_with("cubic-bezier(") => {}
            _ => {
                if let Ok(number) = lower.parse::<f32>() {
                    request.repeats = Some(number.ceil().max(1.0) as u32);
                } else if request.name.is_empty() {
                    request.name = token.to_string();
                }
            }
        }
    }

    request
}

/// How many times an animation repeats: a count, or `infinite` for none.
fn parse_iterations(value: &str) -> Option<Option<u32>> {
    let value = value.trim();

    if value.eq_ignore_ascii_case("infinite") {
        return Some(None);
    }

    value
        .parse::<f32>()
        .ok()
        .map(|number| Some(number.ceil().max(1.0) as u32))
}

/// The animations one element is given by a block of declarations: each named
/// `@keyframes` block, once per property it sets.
fn element_animations(
    element: Node,
    declarations: &[(String, String)],
    stylesheet: &Stylesheet,
    viewport: (f32, f32),
) -> Vec<Animation> {
    let requests = animation_requests(declarations);
    if requests.is_empty() {
        return Vec::new();
    }

    let origin = transform_origin(declarations, element, viewport);
    let mut animations = Vec::new();

    for request in requests {
        let Some(frames) = stylesheet.keyframes.get(&request.name) else {
            continue;
        };

        animations.extend(keyframe_animations(element, frames, &request, origin));
    }

    animations
}

/// The animations an `@keyframes` block gives an element: one per property it sets,
/// with the offsets of the block as their key times.
fn keyframe_animations(
    element: Node,
    frames: &[(f32, String)],
    request: &AnimationRequest,
    origin: Option<(f32, f32)>,
) -> Vec<Animation> {
    // The properties in the order the block first names them, each with what it is at
    // every offset that sets it.
    let mut properties: Vec<(String, Vec<(f32, String)>)> = Vec::new();

    for (offset, body) in frames {
        for (property, value) in declarations(body) {
            if !KEYFRAME_PROPERTIES.contains(&property.as_str()) {
                continue;
            }

            match properties.iter_mut().find(|(named, _)| *named == property) {
                Some((_, values)) => values.push((*offset, value)),
                None => properties.push((property, vec![(*offset, value)])),
            }
        }
    }

    let timing = Timing {
        begin_seconds: request.delay,
        duration_seconds: request.duration.max(0.001),
        repeats: request.repeats,
        freeze: request.freeze,
    };
    let mut animations = Vec::new();

    for (property, mut values) in properties {
        // A block that names only one end of an animation starts or ends at the value
        // the element already carries, which is how a browser reads it: `to { ... }`
        // alone is an animation from what is there to what it says.
        if let Some(base) = base_value(element, &property) {
            if values.first().is_some_and(|(offset, _)| *offset > 0.0) {
                values.insert(0, (0.0, base.clone()));
            }
            if values.last().is_some_and(|(offset, _)| *offset < 1.0) {
                values.push((1.0, base));
            }
        }

        if values.len() < 2 {
            continue;
        }

        let key_times: Vec<f32> = values.iter().map(|(offset, _)| *offset).collect();

        match property.as_str() {
            "transform" => {
                let Some((functions, numbers)) = transform_values(&values) else {
                    continue;
                };

                animations.push(Animation {
                    target: element.id(),
                    kind: Kind::TransformList { functions, origin },
                    timing: Timing { ..timing },
                    key_times,
                    values: Values::Numbers(numbers),
                });
            }
            "fill" | "stroke" => {
                let Some(colors) = values
                    .iter()
                    .map(|(_, value)| parse_color(value))
                    .collect::<Option<Vec<[f32; 3]>>>()
                else {
                    continue;
                };

                animations.push(Animation {
                    target: element.id(),
                    kind: Kind::Attribute(property),
                    timing: Timing { ..timing },
                    key_times,
                    values: Values::Colors(colors),
                });
            }
            name => {
                let Some(numbers) = values
                    .iter()
                    .map(|(_, value)| parse_css_number(value))
                    .collect::<Option<Vec<f32>>>()
                else {
                    continue;
                };

                animations.push(Animation {
                    target: element.id(),
                    kind: Kind::Attribute(name.to_string()),
                    timing: Timing { ..timing },
                    key_times,
                    values: Values::Numbers(numbers.iter().map(|number| vec![*number]).collect()),
                });
            }
        }
    }

    animations
}

/// The value a property has before an animation of it runs: what the element itself
/// says for it, or the value the property starts out as. A property with no such value
/// — a `fill` nothing declares, which is black by the specification — is one this
/// leaves out rather than guesses at.
fn base_value(element: Node, property: &str) -> Option<String> {
    if let Some(value) = element.attribute(property) {
        return Some(value.to_string());
    }

    match property {
        "transform" => Some("none".to_string()),
        "opacity" | "fill-opacity" | "stroke-opacity" => Some("1".to_string()),
        _ => None,
    }
}

/// The functions one `transform` keyframe names and the numbers between them: the
/// arguments of every function of every keyframe, flattened, in the one order the
/// functions are written in. A keyframe that says `none` counts as the identity of the
/// others, which is what makes `from { transform: none }` a rotation from nothing.
fn transform_values(values: &[(f32, String)]) -> Option<(Vec<TransformFunction>, Vec<Vec<f32>>)> {
    let parsed: Vec<Vec<(TransformFunction, Vec<f32>)>> = values
        .iter()
        .map(|(_, value)| parse_css_transform(value))
        .collect::<Option<Vec<_>>>()?;

    let template = parsed.iter().find(|list| !list.is_empty())?.clone();
    let mut numbers = Vec::new();

    for list in &parsed {
        if list.is_empty() {
            numbers.push(
                template
                    .iter()
                    .flat_map(|(function, _)| {
                        function
                            .name
                            .identity()
                            .iter()
                            .take(function.arity)
                            .copied()
                    })
                    .collect(),
            );
            continue;
        }

        if list.len() != template.len()
            || !list
                .iter()
                .zip(template.iter())
                .all(|((function, args), (other, other_args))| {
                    function.name as u8 == other.name as u8
                        && function.arity == other.arity
                        && args.len() == other_args.len()
                })
        {
            return None;
        }

        numbers.push(
            list.iter()
                .flat_map(|(_, arguments)| arguments.iter().copied())
                .collect(),
        );
    }

    Some((
        template.iter().map(|(function, _)| *function).collect(),
        numbers,
    ))
}

/// A CSS transform list, as the functions it names and their arguments.
fn parse_css_transform(value: &str) -> Option<Vec<(TransformFunction, Vec<f32>)>> {
    let value = value.trim();

    if value.is_empty() || value.eq_ignore_ascii_case("none") {
        return Some(Vec::new());
    }

    let mut functions = Vec::new();
    let mut rest = value;

    while let Some(open) = rest.find('(') {
        let name = rest[..open].trim();
        let close = open + rest[open..].find(')')?;
        let kind = css_transform(name)?;

        let arguments: Vec<f32> = rest[open + 1..close]
            .split([',', ' ', '\t'])
            .filter(|argument| !argument.trim().is_empty())
            .map(|argument| css_argument(kind, argument.trim()))
            .collect::<Option<Vec<f32>>>()?;

        if arguments.is_empty() {
            return None;
        }

        functions.push((
            TransformFunction {
                name: kind,
                arity: arguments.len(),
            },
            arguments,
        ));
        rest = &rest[close + 1..];
    }

    Some(functions)
}

fn css_transform(name: &str) -> Option<CssTransform> {
    match name.to_ascii_lowercase().as_str() {
        "translate" | "translatex" | "translatey" => Some(CssTransform::Translate),
        "scale" | "scalex" | "scaley" => Some(CssTransform::Scale),
        "rotate" => Some(CssTransform::Rotate),
        "skewx" => Some(CssTransform::SkewX),
        "skewy" => Some(CssTransform::SkewY),
        "matrix" => Some(CssTransform::Matrix),
        _ => None,
    }
}

/// One argument of a transform function: an angle for the ones that turn, a length for
/// the ones that move. A percentage is not something this can resolve without laying
/// the element out, so it is refused rather than guessed at.
fn css_argument(kind: CssTransform, argument: &str) -> Option<f32> {
    match kind {
        CssTransform::Rotate | CssTransform::SkewX | CssTransform::SkewY => css_angle(argument),
        _ => parse_css_number(argument),
    }
}

/// An angle in degrees, whatever it is written in.
fn css_angle(argument: &str) -> Option<f32> {
    let argument = argument.trim();
    let lower = argument.to_ascii_lowercase();

    if let Some(turns) = lower.strip_suffix("turn") {
        return turns.trim().parse::<f32>().ok().map(|turns| turns * 360.0);
    }

    if let Some(radians) = lower.strip_suffix("rad") {
        return radians
            .trim()
            .parse::<f32>()
            .ok()
            .map(|radians| radians.to_degrees());
    }

    if let Some(gradians) = lower.strip_suffix("grad") {
        return gradians
            .trim()
            .parse::<f32>()
            .ok()
            .map(|gradians| gradians * 0.9);
    }

    if let Some(degrees) = lower.strip_suffix("deg") {
        return degrees.trim().parse().ok();
    }

    lower.parse().ok()
}

/// A number as CSS writes it for an attribute that takes one: a length in pixels, or
/// the number itself.
fn parse_css_number(value: &str) -> Option<f32> {
    let value = value.trim().to_ascii_lowercase();

    match value.strip_suffix("px") {
        Some(number) => number.trim().parse().ok(),
        None => value.parse().ok(),
    }
}

/// A time as CSS writes it, which is `2s` or `500ms`: a bare number is not one, and
/// reading it as one would turn a repeat count into a duration.
fn parse_css_time(value: &str) -> Option<f32> {
    let value = value.trim().to_ascii_lowercase();

    if let Some(milliseconds) = value.strip_suffix("ms") {
        return milliseconds.trim().parse::<f32>().ok();
    }

    value.strip_suffix('s')?.trim().parse::<f32>().ok()
}

/// What a transform is turned around, resolved into user units: CSS's own default is
/// the origin of the coordinate system, so a document that says nothing gets nothing
/// and one that asks for a center gets the center of the box it draws in.
fn transform_origin(
    declarations: &[(String, String)],
    element: Node,
    viewport: (f32, f32),
) -> Option<(f32, f32)> {
    let value = declarations
        .iter()
        .find(|(property, _)| property == "transform-origin")
        .map(|(_, value)| value.trim().to_string())?;
    let fill_box = declarations.iter().any(|(property, value)| {
        property == "transform-box" && value.trim().eq_ignore_ascii_case("fill-box")
    });

    // A percentage is of the element's own box under `fill-box` and of the viewport
    // otherwise, which is the difference the property exists to make.
    let box_size = if fill_box {
        element_box(element)
            .map(|(left, top, right, bottom)| {
                ((right - left).max(1.0), (bottom - top).max(1.0), left, top)
            })
            .unwrap_or((viewport.0, viewport.1, 0.0, 0.0))
    } else {
        (viewport.0, viewport.1, 0.0, 0.0)
    };

    let tokens: Vec<&str> = value.split_whitespace().collect();
    if tokens.is_empty() || tokens.len() > 2 {
        return None;
    }

    let axis = |token: &str, index: usize| -> Option<f32> {
        let lower = token.to_ascii_lowercase();
        let percent = |value: f32| match index {
            0 => box_size.2 + box_size.0 * value / 100.0,
            _ => box_size.3 + box_size.1 * value / 100.0,
        };

        match lower.as_str() {
            "left" | "top" => Some(0.0),
            "center" => Some(match index {
                0 => box_size.2 + box_size.0 / 2.0,
                _ => box_size.3 + box_size.1 / 2.0,
            }),
            "right" | "bottom" => Some(match index {
                0 => box_size.2 + box_size.0,
                _ => box_size.3 + box_size.1,
            }),
            _ => {
                if let Some(percentage) = lower.strip_suffix('%') {
                    return percentage.trim().parse::<f32>().ok().map(percent);
                }

                parse_css_number(&lower)
            }
        }
    };

    let (x, y) = match tokens.len() {
        1 => (axis(tokens[0], 0)?, axis(tokens[0], 1)?),
        _ => (axis(tokens[0], 0)?, axis(tokens[1], 1)?),
    };

    (x != 0.0 || y != 0.0).then_some((x, y))
}

/// The box an element draws in, as far as its own markup says: enough to turn CSS's
/// rotation about a center into one the renderer can draw. A shape this cannot measure
/// — a path, a text run — has no box, and what asks for one falls back to the viewport.
fn element_box(element: Node) -> Option<(f32, f32, f32, f32)> {
    let number = |name: &str| {
        element
            .attribute(name)
            .and_then(parse_css_number)
            .unwrap_or(0.0)
    };

    match element.tag_name().name() {
        "rect" => {
            let (x, y) = (number("x"), number("y"));

            Some((x, y, x + number("width"), y + number("height")))
        }
        "circle" => {
            let (x, y, r) = (number("cx"), number("cy"), number("r"));

            Some((x - r, y - r, x + r, y + r))
        }
        "ellipse" => {
            let (x, y, rx, ry) = (number("cx"), number("cy"), number("rx"), number("ry"));

            Some((x - rx, y - ry, x + rx, y + ry))
        }
        "line" => Some((number("x1"), number("y1"), number("x2"), number("y2"))),
        "polyline" | "polygon" => {
            let points = element.attribute("points").unwrap_or("");
            let numbers: Vec<f32> = points
                .split([',', ' ', '\t', '\n'])
                .filter(|part| !part.trim().is_empty())
                .map(|part| part.trim().parse::<f32>())
                .collect::<Result<Vec<f32>, _>>()
                .ok()?;

            let (mut left, mut top, mut right, mut bottom) =
                (f32::MAX, f32::MAX, f32::MIN, f32::MIN);
            for pair in numbers.chunks(2) {
                let (x, y) = (pair[0], *pair.get(1).unwrap_or(&pair[0]));
                left = left.min(x);
                right = right.max(x);
                top = top.min(y);
                bottom = bottom.max(y);
            }

            (left <= right).then_some((left, top, right, bottom))
        }
        // A container is as large as everything under it.
        "svg" | "g" | "a" | "switch" => {
            let boxes: Vec<(f32, f32, f32, f32)> = element
                .children()
                .filter(Node::is_element)
                .filter_map(element_box)
                .collect();
            let first = boxes.first()?;

            Some(
                boxes
                    .iter()
                    .fold(*first, |(left, top, right, bottom), box_| {
                        (
                            left.min(box_.0),
                            top.min(box_.1),
                            right.max(box_.2),
                            bottom.max(box_.3),
                        )
                    }),
            )
        }
        _ => None,
    }
}

/// The viewport a document draws in, which is what a percentage origin is of when the
/// element does not ask for its own box.
fn viewport_size(document: &Document) -> (f32, f32) {
    let Some(root) = document.root().first_element_child() else {
        return (300.0, 150.0);
    };

    let width = root.attribute("width").and_then(parse_css_number);
    let height = root.attribute("height").and_then(parse_css_number);

    if let (Some(width), Some(height)) = (width, height) {
        if width > 0.0 && height > 0.0 {
            return (width, height);
        }
    }

    if let Some(view_box) = root.attribute("viewBox") {
        let numbers: Vec<f32> = view_box
            .split([',', ' ', '\t'])
            .filter(|part| !part.trim().is_empty())
            .filter_map(|part| part.trim().parse::<f32>().ok())
            .collect();

        if numbers.len() == 4 && numbers[2] > 0.0 && numbers[3] > 0.0 {
            return (numbers[2], numbers[3]);
        }
    }

    (300.0, 150.0)
}

/// A stylesheet's rule or keyframe blocks, each with the prelude that names it. Braces
/// are counted rather than matched, so a block that holds blocks — which is what
/// `@keyframes` is — comes back whole.
fn blocks(text: &str) -> Vec<(String, String)> {
    let mut blocks = Vec::new();
    let bytes = text.as_bytes();
    let mut index = 0;

    while let Some(open) = text[index..].find('{') {
        let open = index + open;
        let prelude = text[index..open].trim().to_string();

        let mut depth = 1usize;
        let mut cursor = open + 1;
        while cursor < bytes.len() && depth > 0 {
            match bytes[cursor] {
                b'{' => depth += 1,
                b'}' => depth -= 1,
                _ => {}
            }
            cursor += 1;
        }

        let end = cursor.saturating_sub(1).max(open + 1);
        let body = text[open + 1..end].to_string();

        if !prelude.is_empty() {
            blocks.push((prelude, body));
        }

        index = cursor.max(open + 1);
    }

    blocks
}

/// The declarations of a block: property and value, in the order they are written.
fn declarations(text: &str) -> Vec<(String, String)> {
    text.split(';')
        .filter_map(|declaration| {
            let (property, value) = declaration.split_once(':')?;
            let property = property.trim().to_ascii_lowercase();

            (!property.is_empty()).then(|| (property, value.trim().to_string()))
        })
        .collect()
}

fn strip_comments(css: &str) -> String {
    let mut out = String::with_capacity(css.len());
    let mut rest = css;

    while let Some(start) = rest.find("/*") {
        out.push_str(&rest[..start]);
        match rest[start + 2..].find("*/") {
            Some(end) => rest = &rest[start + 2 + end + 2..],
            None => return out,
        }
    }

    out.push_str(rest);
    out
}

/// Whether an element is what a selector names: a comma-separated list of complex
/// selectors, each a chain of type, class and id parts joined by descendant and child
/// combinators. Anything else a stylesheet may hold — attribute selectors,
/// pseudo-classes, `*` — matches nothing rather than something wrong.
fn selector_matches(selector: &str, element: Node) -> bool {
    selector
        .split(',')
        .any(|alternative| complex_matches(alternative.trim(), element))
}

fn complex_matches(selector: &str, element: Node) -> bool {
    let tokens: Vec<String> = selector
        .replace('>', " > ")
        .split_whitespace()
        .map(str::to_string)
        .collect();

    if tokens.is_empty() || tokens[tokens.len() - 1] == ">" {
        return false;
    }

    let mut index = tokens.len() - 1;
    if !compound_matches(&tokens[index], element) {
        return false;
    }

    let mut current = element;
    while index > 0 {
        let child_only = tokens[index - 1] == ">";
        if child_only {
            index -= 1;
        }
        if index == 0 {
            return false;
        }
        index -= 1;

        let compound = &tokens[index];
        if compound == ">" {
            return false;
        }

        let mut ancestor = current.parent_element();
        let mut found = false;

        while let Some(node) = ancestor {
            if compound_matches(compound, node) {
                current = node;
                found = true;
                break;
            }

            if child_only {
                break;
            }

            ancestor = node.parent_element();
        }

        if !found {
            return false;
        }
    }

    true
}

/// Whether an element carries every part of one compound selector, `tag.class#id`.
fn compound_matches(compound: &str, element: Node) -> bool {
    if !element.is_element() {
        return false;
    }

    let mut tag: Option<&str> = None;
    let mut class: Option<&str> = None;
    let mut id: Option<&str> = None;

    let mut rest = compound;
    if let Some(index) = rest.find(['.', '#']) {
        if index > 0 {
            tag = Some(&rest[..index]);
        }
        rest = &rest[index..];
    } else {
        tag = Some(rest);
        rest = "";
    }

    while !rest.is_empty() {
        let marker = rest.as_bytes()[0] as char;
        let end = rest[1..]
            .find(['.', '#'])
            .map(|index| index + 1)
            .unwrap_or(rest.len());
        let name = &rest[1..end];

        if name.is_empty() {
            return false;
        }

        match marker {
            '.' if class.is_none() => class = Some(name),
            '#' if id.is_none() => id = Some(name),
            _ => return false,
        }

        rest = &rest[end..];
    }

    if let Some(tag) = tag {
        if !tag.is_empty() && element.tag_name().name() != tag {
            return false;
        }
    }

    if let Some(class) = class {
        let classes = element.attribute("class").unwrap_or("");

        if !classes.split_whitespace().any(|name| name == class) {
            return false;
        }
    }

    if let Some(id) = id {
        if element.attribute("id") != Some(id) {
            return false;
        }
    }

    true
}

/// The transform list as the renderer reads it: the functions in the order they were
/// written, with their arguments blended, and the origin written around them so that
/// what turns or scales turns or scales about the point CSS named.
fn write_transform_list(
    functions: &[TransformFunction],
    origin: Option<(f32, f32)>,
    numbers: &[f32],
) -> String {
    let mut text = String::new();
    let mut index = 0;

    if let Some((x, y)) = origin {
        text.push_str(&format!(
            "translate({} {}) ",
            format_number(x),
            format_number(y)
        ));
    }

    for function in functions {
        let end = (index + function.arity).min(numbers.len());
        text.push_str(function.name.as_str());
        text.push('(');

        for (position, number) in numbers[index.min(end)..end].iter().enumerate() {
            if position > 0 {
                text.push(' ');
            }
            text.push_str(&format_number(*number));
        }

        text.push_str(") ");
        index = end;
    }

    if let Some((x, y)) = origin {
        text.push_str(&format!(
            "translate({} {})",
            format_number(-x),
            format_number(-y)
        ));
    }

    text.trim_end().to_string()
}

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

    /// A declaration kept apart from what it moves names it: a clock's hands are at
    /// the end of the file and the animations that turn them are in `<defs>`, pointing
    /// at each hand by id.
    #[test]
    fn animates_the_element_a_declaration_names() {
        for href in [r##"xlink:href="#hand""##, r##"href="#hand""##] {
            let xml = format!(
                r##"<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="100" height="100"><defs><animateTransform {href} attributeName="transform" type="rotate" from="0 50 50" to="360 50 50" dur="1s" repeatCount="indefinite"/></defs><path id="hand" d="M50 50 L50 10" transform="rotate(90 50 50)"/></svg>"##
            );
            let (document, playback) = playback(&xml);

            let first = playback.document_at(&document, 0);
            assert!(
                first.contains(r#"transform="rotate(90 50 50) rotate(0 50 50)""#),
                "{first}"
            );

            let middle = playback.document_at(&document, 15);
            assert!(middle.contains("rotate(178.2 50 50)"), "{middle}");
        }
    }

    /// A declaration that names an element which is not there is about the element it
    /// sits in, which is what a document that never says `href` means.
    #[test]
    fn falls_back_to_the_element_a_declaration_sits_in() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"><animate attributeName="opacity" from="0" to="1" dur="1s" repeatCount="indefinite"/></rect></svg>"#,
        );

        assert!(playback
            .document_at(&document, 15)
            .contains(r#"opacity="0.495""#));
    }

    /// A pass that holds the whole of what a document says is repeated, and one that
    /// was cut short is played forward instead: a clock's second hand keeps going.
    #[test]
    fn a_pass_that_was_cut_short_is_not_repeated() {
        let (_, quick) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"><animate attributeName="opacity" from="0" to="1" dur="2s" repeatCount="indefinite"/></rect></svg>"#,
        );
        assert!(quick.repeats_its_pass());

        let (_, slow) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"><animateTransform attributeName="transform" type="rotate" from="0" to="360" dur="43200s" repeatCount="indefinite"/></rect></svg>"#,
        );
        assert_eq!(slow.frames(), MAX_FRAMES);
        assert!(!slow.repeats_its_pass());
    }

    /// A stylesheet's `@keyframes`, named by a rule, are played the same way a
    /// declaration in the markup is.
    #[test]
    fn plays_a_keyframes_animation_a_stylesheet_names() {
        let (document, playback) = playback(
            r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><style>
                 @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
                 .spinner { animation: spin 1s linear infinite; }
               </style><rect class="spinner" x="40" y="40" width="20" height="20" fill="#2b5fd9"/></svg>"##,
        );

        let first = playback.document_at(&document, 0);
        assert!(first.contains(r#"transform="rotate(0)""#), "{first}");

        // 495 milliseconds into a one-second pass is 178.2 degrees.
        let middle = playback.document_at(&document, 15);
        assert!(middle.contains(r#"transform="rotate(178.2)""#), "{middle}");
        assert!(
            first.contains(r#"class="spinner""#),
            "the document is written as it was found: {first}"
        );
    }

    /// A keyframe that says `transform: none` is the identity of the ones that do not,
    /// which is what makes a rotation from nothing play.
    #[test]
    fn reads_a_keyframe_that_names_no_transform_as_the_identity() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><style>
                 @keyframes turn { from { transform: none; } to { transform: rotate(90deg); } }
                 rect { animation: turn 1s linear infinite; }
               </style><rect x="10" y="10" width="20" height="20"/></svg>"#,
        );

        let first = playback.document_at(&document, 0);
        assert!(first.contains(r#"transform="rotate(0)""#), "{first}");

        let last = playback.document_at(&document, 15);
        assert!(last.contains(r#"transform="rotate(44.55)""#), "{last}");
    }

    /// A percentage origin is of the element's own box under `fill-box`, and what that
    /// buys is a rotation about the middle of the shape rather than the corner.
    #[test]
    fn turns_about_the_origin_a_stylesheet_names() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><style>
                 @keyframes spin { to { transform: rotate(180deg); } }
                 rect { animation: spin 1s linear infinite; transform-origin: center; transform-box: fill-box; }
               </style><rect x="40" y="40" width="20" height="20"/></svg>"#,
        );

        let half = playback.document_at(&document, 15);
        assert!(
            half.contains(r#"transform="translate(50 50) rotate(89.1) translate(-50 -50)""#),
            "{half}"
        );
    }

    #[test]
    fn plays_a_keyframes_animation_an_element_names_in_its_own_style() {
        let (document, playback) = playback(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><style>
                 @keyframes fade { 0% { opacity: 1; } 50% { opacity: 0.25; } 100% { opacity: 1; } }
               </style><rect width="10" height="10" style="animation: fade 2s linear infinite"/></svg>"#,
        );

        assert!(playback
            .document_at(&document, 0)
            .contains(r#"opacity="1""#));

        // 495 milliseconds into a two-second pass is 495 thousandths of the way to the
        // middle value of 0.25.
        let quarter = playback.document_at(&document, 15);
        assert!(quarter.contains(r#"opacity="0.629""#), "{quarter}");
    }

    /// What is written is a document the renderer can draw, and two frames of an
    /// animation are two different pictures: the whole path a played frame takes, from
    /// the declarations to the pixels.
    #[test]
    fn draws_frames_that_move() {
        let (document, playback) = playback(
            r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><style>
                 @keyframes spin { to { transform: rotate(180deg); } }
                 rect { animation: spin 1s linear infinite; transform-origin: center; transform-box: fill-box; }
               </style><rect x="30" y="10" width="40" height="80" fill="#2b5fd9"/></svg>"##,
        );

        let first = crate::svg_preview::render_text(&playback.document_at(&document, 0), 100, 100)
            .expect("a drawn frame");
        let later = crate::svg_preview::render_text(&playback.document_at(&document, 15), 100, 100)
            .expect("a drawn frame");

        assert_eq!((first.1, first.2), (later.1, later.2), "the same box");
        assert_ne!(first.0, later.0, "the frame moved");
    }

    /// A stylesheet that only paints — no `@keyframes` — leaves a document that stands
    /// still, so nothing is played for it.
    #[test]
    fn a_stylesheet_without_keyframes_plays_nothing() {
        let document = Document::parse(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><style>.a { fill: #ff0000; }</style><rect class="a" width="10" height="10"/></svg>"#,
        )
        .expect("a parsed document");

        assert!(Playback::parse(&document).is_none());
    }

    /// The selectors a hover's documents use are read the way a browser reads them,
    /// and the ones this cannot match match nothing rather than something wrong.
    #[test]
    fn matches_the_selectors_it_knows() {
        let document = Document::parse(
            r#"<svg xmlns="http://www.w3.org/2000/svg"><g class="icon"><rect id="bar" class="a b" width="10" height="10"/></g></svg>"#,
        )
        .expect("a parsed document");
        let rect = document
            .descendants()
            .find(|node| node.has_tag_name("rect"))
            .expect("the rect");

        for selector in [
            "rect",
            ".a",
            "#bar",
            "rect.a",
            "rect#bar.a",
            ".icon rect",
            ".icon > rect",
            "g > rect",
            "svg .icon rect",
            ".a, .b",
            "rect, circle",
        ] {
            assert!(selector_matches(selector, rect), "{selector}");
        }

        for selector in [
            "circle",
            ".b.a", // both classes are on the element, but a compound names each once
            "svg > rect",
            "rect > rect",
            "[class]",
            ":first-child",
            "*",
        ] {
            assert!(!selector_matches(selector, rect), "{selector}");
        }
    }
}
