//! What a picture whose samples are light is shown as.
//!
//! An `.exr`, a Radiance `.hdr` and a float texture do not hold levels — the numbers a
//! screen is addressed with — but light: `1.0` is as white as the display can draw, `0.18`
//! is a mid grey, and a window or a lamp is tens of times either. What a preview is
//! composed in is eight bits to the channel and already display-encoded, so a picture of
//! that kind needs two things a PNG does not: the transfer function that says how a level
//! is written down, and a curve that brings a range wider than the display into it rather
//! than clipping what is above it.
//!
//! Measured rather than assumed, what happens without them — the `image` crate's own
//! conversion is a bare clamp with no transfer at all — is that a linear `0.5`, which is a
//! light grey and is `188` as a level, is drawn as `128`, and everything from `1.0` up is
//! drawn as white. A preview of a render or a photograph comes out dark in the midtones
//! with its highlights burnt out, and the worse the picture's range the worse it reads.
//!
//! What is applied here instead is an exposure, a tone curve and the sRGB transfer, in
//! that order, which is the shape every viewer of a photograph or a render takes. The
//! transfer is what makes the numbers levels; the curve is what keeps a value ten times
//! white from being ten times clipped — Reinhard's `x / (1 + x)` by default, which maps
//! the whole of the positive range into the display's and leaves the values a screen can
//! already show very nearly where they were. Both are settings: `hdr_tone_map` in
//! `config.ini` names the curve — `reinhard` (the default), `srgb`, `aces`, or `off` for
//! the clamp the app drew before this module existed — and `hdr_exposure` shifts the
//! picture in stops before either is applied.
//!
//! Nothing here allocates for a picture that is not of that kind, and nothing is applied
//! to one: a PNG, a JPEG or a BMP holds levels already, and a level put through a transfer
//! function a second time is a washed-out picture rather than a corrected one.

/// How a picture whose samples are light is brought down to eight bits.
#[derive(Clone, Copy)]
pub struct ToneMap {
    pub curve: Curve,
    /// Stops the picture is shifted by before the curve: `0` is as the file holds it,
    /// `+1` is twice as bright, `-1` half as bright.
    pub exposure: f32,
}

/// The curve a picture's range is brought into the display's with.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Curve {
    /// No transfer and no curve: what a sample says is what is drawn, clamped. This is
    /// the answer the app gave before this module existed, kept for the picture that is
    /// already graded for a screen.
    Off,
    /// The sRGB transfer alone. Correct for a picture whose range is already within the
    /// display's, and a clip for one whose is not.
    Srgb,
    /// Reinhard's `x / (1 + x)`, then the transfer: the whole of the positive range is
    /// brought into `0..1`, nothing is clipped, and a value the display can already show
    /// is left very nearly where it was.
    Reinhard,
    /// The ACES filmic curve, then the transfer: darker in the shadows and more
    /// saturated than Reinhard, which is what a filmic answer looks like.
    Aces,
}

/// The curve a file names, which is the setting's own spelling.
impl Curve {
    pub fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "off" | "none" | "clamp" => Some(Curve::Off),
            "srgb" => Some(Curve::Srgb),
            "reinhard" => Some(Curve::Reinhard),
            "aces" => Some(Curve::Aces),
            _ => None,
        }
    }

    /// The spelling a value is written back out with.
    pub fn as_str(self) -> &'static str {
        match self {
            Curve::Off => "off",
            Curve::Srgb => "srgb",
            Curve::Reinhard => "reinhard",
            Curve::Aces => "aces",
        }
    }
}

impl ToneMap {
    /// What the configuration says, read each time rather than captured, so an edit
    /// applies to the next hover without a restart — the way every other setting here is
    /// read.
    pub fn current() -> Self {
        crate::CONFIG
            .lock()
            .map(|config| Self {
                curve: config.hdr_tone_map,
                exposure: crate::config::config::sanitize_hdr_exposure(config.hdr_exposure),
            })
            .unwrap_or(Self {
                curve: crate::config::config::DEFAULT_HDR_TONE_MAP,
                exposure: 0.0,
            })
    }

    /// One sample of light as a level.
    pub fn encode(&self, linear: f32) -> u8 {
        let exposed = if self.exposure == 0.0 {
            linear
        } else {
            linear * 2.0f32.powf(self.exposure)
        };

        // A sample that is not a number — a half float that carries an infinity, a
        // channel an encoder wrote garbage into — is a level of nothing rather than a
        // value that propagates through the rest of the frame.
        let exposed = if exposed.is_finite() { exposed } else { 0.0 };

        let curved = match self.curve {
            // Neither of these two has a curve: one is the transfer and nothing else, and
            // the other is nothing at all.
            Curve::Off | Curve::Srgb => exposed,
            Curve::Reinhard => {
                let positive = exposed.max(0.0);
                positive / (1.0 + positive)
            }
            Curve::Aces => {
                const A: f32 = 2.51;
                const B: f32 = 0.03;
                const C: f32 = 2.43;
                const D: f32 = 0.59;
                const E: f32 = 0.14;

                let positive = exposed.max(0.0);
                ((positive * (A * positive + B)) / (positive * (C * positive + D) + E))
                    .clamp(0.0, 1.0)
            }
        };

        let level = match self.curve {
            Curve::Off => curved.clamp(0.0, 1.0),
            _ => encode_srgb(curved.clamp(0.0, 1.0)),
        };

        (level * 255.0 + 0.5) as u8
    }
}

/// The sRGB transfer function: the curve that says how a level is written down for a
/// display. Its two pieces are the standard's own — a linear segment near black, which is
/// where the curve's slope would run away, and a power above it.
fn encode_srgb(linear: f32) -> f32 {
    if linear <= 0.003_130_8 {
        linear * 12.92
    } else {
        1.055 * linear.powf(1.0 / 2.4) - 0.055
    }
}
