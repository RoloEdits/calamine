use thiserror::Error;

#[derive(Debug, Error)]
pub enum XlsxError {
    #[error("failed to read file {0}")]
    ReadFailure(#[from] std::io::Error),
    #[error("failed to open zip archive from path: {0}")]
    ArchiveFailure(#[from] zip::result::ZipError),
    #[error("end of file error `{0}`")]
    XmlEof(&'static str),
    #[error("end of file error `{0}`")]
    NumberParseError(#[from] std::num::ParseIntError),
}

#[derive(Debug, Error)]
pub enum ParseError {
    #[error("invalid xml {0}")]
    InvalidXML(String),
}
