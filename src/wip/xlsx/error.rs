use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("failed to read file {0}")]
    ReadFailure(#[from] std::io::Error),
    #[error("failed to open zip archive from path: {0}")]
    ArchiveFailure(#[from] zip::result::ZipError),
    #[error("end of file error `{0}`")]
    XmlEof(&'static str),
    #[error("failed to parse number `{0}`")]
    NumberParseError(#[from] std::num::ParseIntError),
    #[error("xml parsng error `{0}`")]
    XmlParsingError(#[from] quick_xml::Error),
}
