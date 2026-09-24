// The settings module is reached as `config::config` from every reader and engine,
// which is the path this app has always named it by: renaming it would touch a
// hundred import sites to answer a lint about the name itself.
#[allow(clippy::module_inception)]
pub(crate) mod config;
pub(crate) mod theme_files;
