use std::borrow::Cow;

use compact_str::{CompactString, ToCompactString};

use super::style::Font;

#[derive(Debug, Clone, Default)]
pub struct Cell<'a> {
    pub(crate) value: Option<Cow<'a, CompactString>>,
    pub(crate) r#type: Option<Type>,
    pub(crate) column: u32,
    pub(crate) row: u32,
    pub(crate) font: Cow<'a, Font>,
}

impl Cell<'_> {
    #[must_use]
    pub fn new(column: u32, row: u32) -> Self {
        Self {
            value: None,
            column,
            row,
            font: Cow::Owned(Font::default()),
            r#type: None,
        }
    }

    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        match self.value.as_ref() {
            Some(value) => Some(value.as_str()),
            None => None,
        }
    }

    pub fn insert_value<V: IntoCellValue>(&mut self, value: V) {
        let (value, r#type) = value.into();
        self.r#type = Some(r#type);
        self.value = Some(Cow::Owned(value));
    }

    #[inline]
    #[must_use]
    pub fn font(&self) -> &Font {
        self.font.as_ref()
    }
}

pub trait IntoCellValue {
    fn into(self) -> (CompactString, Type);
}

impl IntoCellValue for i32 {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for u32 {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for i64 {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for u64 {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for usize {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for isize {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for f32 {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for f64 {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::Number)
    }
}

impl IntoCellValue for &str {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::String)
    }
}

impl IntoCellValue for String {
    fn into(self) -> (CompactString, Type) {
        (self.to_compact_string(), Type::String)
    }
}

#[derive(Debug, Clone, Copy)]
pub enum Type {
    Number,
    String,
    Formula,
}
