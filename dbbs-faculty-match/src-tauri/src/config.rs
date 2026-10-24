pub(crate) const FACULTY_DATASET_BASENAME: &str = "faculty_dataset";
pub(crate) const FACULTY_DATASET_DEFAULT_EXTENSION: &str = "tsv";
pub(crate) const FACULTY_DATASET_EXTENSIONS: &[&str] = &["tsv", "txt", "xlsx", "xls"];
pub(crate) const DEFAULT_FACULTY_DATASET: &[u8] =
    include_bytes!("../assets/default_faculty_dataset.tsv");
pub(crate) const FACULTY_DATASET_METADATA_NAME: &str = "faculty_dataset_metadata.json";
pub(crate) const FACULTY_DATASET_SOURCE_NAME: &str = "faculty_dataset_source.txt";
pub(crate) const FACULTY_EMBEDDINGS_NAME: &str = "faculty_embeddings.json";
pub(crate) const DEFAULT_FACULTY_EMBEDDINGS: &[u8] =
    include_bytes!("../assets/default_faculty_embeddings.json");
pub(crate) const DEFAULT_EMBEDDING_MODEL: &str = "NeuML/pubmedbert-base-embeddings";
pub(crate) const FACULTY_EMBEDDING_PROGRESS_EVENT: &str = "faculty-embedding-progress";
pub(crate) const FACULTY_EMBEDDING_STANDBY_MESSAGE: &str = "Embedding helper ready.";
