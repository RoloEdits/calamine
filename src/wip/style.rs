use compact_str::CompactString;

#[derive(Debug, Default, Clone)]
pub struct Font {
    pub(crate) name: CompactString,
    pub(crate) size: f64,
    pub(crate) color: CompactString,
}

impl Font {
    #[inline]
    #[must_use]
    pub fn rgb(&self) -> &str {
        &self.color[2..]
    }
    #[inline]
    #[must_use]
    pub fn argb(&self) -> &str {
        &self.color
    }
}

impl AsRef<Font> for Font {
    fn as_ref(&self) -> &Font {
        self
    }
}
