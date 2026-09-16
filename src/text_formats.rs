use crate::CONFIG;
use std::path::Path;

/// Extensions previewed as text before the list is edited in `config.ini`.
///
/// The list is written to the configuration on first run and read back from
/// there, so adding or removing an extension is an edit in the file rather than
/// a rebuild. Anything in it that no syntax definition claims is still
/// previewed, as plain text.
///
/// `.ts` and `.mts` belong to both lists: they are TypeScript sources and MPEG
/// transport streams, so the video gate decides those two by content (an MPEG-TS
/// sync byte) and a TypeScript file falls through to a text preview. Every other
/// extension here is disjoint from the image, video and PDF gates.
pub const DEFAULT_TEXT_EXTENSIONS: &str = "txt,text,log,nfo,diz,md,markdown,mkd,mdown,rst,adoc,asciidoc,org,tex,latex,bib,man,csv,tsv,srt,vtt,rtf,diff,patch,\
ini,cfg,conf,properties,env,toml,yaml,yml,json,jsonc,json5,jsonl,xml,xsd,xsl,xslt,plist,\
html,htm,xhtml,css,scss,sass,less,styl,js,mjs,cjs,jsx,ts,tsx,mts,cts,vue,svelte,astro,\
php,phtml,asp,aspx,jsp,cshtml,erb,haml,slim,ejs,hbs,mustache,twig,liquid,njk,\
sh,bash,zsh,fish,ksh,csh,bat,cmd,ps1,psm1,psd1,vbs,awk,tcl,pl,pm,py,pyw,pyi,rb,rake,lua,\
groovy,gradle,r,rmd,jl,c,h,cc,cpp,cxx,hh,hpp,hxx,cs,csx,java,kt,kts,scala,go,rs,swift,\
m,mm,dart,zig,nim,d,pas,f,f77,f90,f95,f03,for,ftn,adb,ads,hs,lhs,ml,mli,fs,fsi,fsx,ex,exs,\
erl,hrl,elm,clj,cljs,cljc,edn,lisp,lsp,el,scm,ss,rkt,v,sv,svh,vhd,vhdl,sol,sql,graphql,gql,\
proto,cmake,mk,mak,ninja,bzl,tf,tfvars,hcl,nix,asm,s,inc,nasm,wat,ll,glsl,vert,frag,geom,\
comp,hlsl,fx,cg,wgsl,metal,ipynb,lock";

/// Read one extension out of the configured list into the lowercase form the
/// lookups use. A leading dot is accepted because `py` and `.py` are both what
/// a user might type, and an entry that is not a bare extension is dropped so a
/// stray path or sentence in the list cannot turn into a match.
fn sanitized_extension(value: &str) -> Option<String> {
    let trimmed = value.trim().trim_start_matches('.').to_lowercase();
    if trimmed.is_empty() {
        return None;
    }

    let is_extension = trimmed
        .chars()
        .all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '_' | '#'));
    is_extension.then_some(trimmed)
}

/// File names previewed as text before the list is edited in `config.ini`.
///
/// An extension is not enough for the files a repository is recognized by: a
/// `.gitignore` has no extension at all as far as the path is concerned — the dot
/// is the start of its *name* — and `LICENSE`, `Makefile` and `Dockerfile` have no
/// dot in them anywhere. So the gate has a second list, of names, and a file
/// matches if either list does.
pub const DEFAULT_TEXT_NAMES: &str = "license,licence,unlicense,copying,copyright,notice,authors,contributors,contributing,code_of_conduct,security,changelog,changes,history,install,readme,makefile,gnumakefile,makefile.am,makefile.in,dockerfile,containerfile,vagrantfile,gemfile,rakefile,procfile,brewfile,jenkinsfile,justfile,caddyfile,cmakelists.txt,\
.gitignore,.gitattributes,.gitmodules,.gitkeep,.gitconfig,.mailmap,.dockerignore,.editorconfig,.npmignore,.eslintignore,.prettierignore,.babelrc,.eslintrc,.prettierrc,.stylelintrc,.htaccess,.env,.env.local,.env.example,.clang-format,.clang-tidy,.rustfmt.toml,.golangci.yml";

/// Read one name out of the configured list into the lowercase form the lookups
/// use. A leading dot is accepted and dropped, because `.gitignore` and
/// `gitignore` are the same file to everything except the filesystem, and an entry
/// that is not a plausible file name is dropped so a stray path or sentence in the
/// list cannot turn into a match.
fn sanitized_name(value: &str) -> Option<String> {
    let trimmed = value.trim().trim_start_matches('.').to_lowercase();
    if trimmed.is_empty() {
        return None;
    }

    let is_name = trimmed
        .chars()
        .all(|c| c.is_ascii_alphanumeric() || matches!(c, '.' | '-' | '_'));
    is_name.then_some(trimmed)
}

/// The configured name list split into the entries lookups compare against.
pub fn sanitize_names(list: &str) -> Vec<String> {
    let mut names: Vec<String> = Vec::new();
    for entry in list.split(',') {
        if let Some(name) = sanitized_name(entry) {
            if !names.contains(&name) {
                names.push(name);
            }
        }
    }
    names
}

/// The name `path` will be looked up by: its file name in lowercase, with a
/// leading dot dropped, so `.gitignore` and `gitignore` are one entry.
fn lookup_name(path: &Path) -> Option<String> {
    let name = path.file_name()?.to_str()?.to_lowercase();
    if name.is_empty() {
        return None;
    }
    Some(name.trim_start_matches('.').to_string())
}

/// The configured list split into the entries lookups compare against.
pub fn sanitize_extensions(list: &str) -> Vec<String> {
    let mut extensions: Vec<String> = Vec::new();
    for entry in list.split(',') {
        if let Some(extension) = sanitized_extension(entry) {
            if !extensions.contains(&extension) {
                extensions.push(extension);
            }
        }
    }
    extensions
}

/// The extension `path` will be looked up by.
///
/// A dot file is the reason this is not one call to `Path::extension`: `.gitignore`
/// has no extension — the dot begins its name — so a leading dot is stripped and
/// what follows is the name the lists are written with, which is what makes
/// `gitignore` in the extension list mean `.gitignore`.
fn lookup_extension(path: &Path) -> Option<String> {
    if let Some(extension) = path.extension().and_then(|ext| ext.to_str()) {
        return Some(extension.to_lowercase());
    }

    let name = path.file_name()?.to_str()?;
    let stripped = name.strip_prefix('.')?;
    (!stripped.is_empty() && !stripped.contains('.')).then(|| stripped.to_lowercase())
}

/// Whether `path` carries an extension the configuration previews as text.
pub fn matches_configured_extension(path: &Path, extensions: &[String]) -> bool {
    let Some(extension) = lookup_extension(path) else {
        return false;
    };

    extensions.contains(&extension)
}

/// Whether `path` carries a name the configuration previews as text.
pub fn matches_configured_name(path: &Path, names: &[String]) -> bool {
    let Some(name) = lookup_name(path) else {
        return false;
    };

    names.contains(&name)
}

/// Whether the file is previewed as text under the current configuration. The
/// `Enable Text Preview` toggle is checked first, so turning text previews off
/// leaves the lists alone and turning them back on restores them.
pub fn is_text_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| {
            config.text_preview_enabled
                && (matches_configured_extension(path, &config.text_extensions)
                    || matches_configured_name(path, &config.text_names))
        })
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::{
        matches_configured_extension, matches_configured_name, sanitize_extensions, sanitize_names,
        DEFAULT_TEXT_EXTENSIONS, DEFAULT_TEXT_NAMES,
    };
    use crate::video_formats::VIDEO_EXTENSIONS;
    use std::path::Path;

    fn default_list() -> Vec<String> {
        sanitize_extensions(DEFAULT_TEXT_EXTENSIONS)
    }

    #[test]
    fn default_extensions_are_lowercase_and_unique() {
        let extensions = default_list();
        let mut sorted = extensions.clone();
        sorted.sort_unstable();
        sorted.dedup();
        assert_eq!(sorted.len(), extensions.len());
        assert!(extensions.iter().all(|ext| *ext == ext.to_lowercase()));
        assert!(!extensions.is_empty());
    }

    /// The two extensions shared with the video gate are the only overlap, and
    /// they are there because the video gate resolves them by content.
    #[test]
    fn default_extensions_avoid_the_video_names_that_are_not_shared() {
        for extension in default_list() {
            if extension == "ts" || extension == "mts" {
                continue;
            }
            assert!(
                !VIDEO_EXTENSIONS.contains(&extension.as_str()),
                "{extension} is also a video extension"
            );
        }
    }

    #[test]
    fn list_edits_are_normalized() {
        assert_eq!(
            sanitize_extensions(" py, .PY ,md,md,, jsonc "),
            vec!["py", "md", "jsonc"]
        );
        // An entry that is not a bare extension is dropped rather than matched.
        assert_eq!(
            sanitize_extensions("c:/temp/file.txt, *.log, ok"),
            vec!["ok"]
        );
        assert!(sanitize_extensions("  ,  ").is_empty());
    }

    #[test]
    fn extension_gate_matches_configured_files() {
        let extensions = sanitize_extensions("txt,py,md");
        assert!(matches_configured_extension(
            Path::new("notes.txt"),
            &extensions
        ));
        assert!(matches_configured_extension(
            Path::new("NOTES.TXT"),
            &extensions
        ));
        assert!(matches_configured_extension(
            Path::new("main.py"),
            &extensions
        ));
        assert!(!matches_configured_extension(
            Path::new("main.rs"),
            &extensions
        ));
        assert!(!matches_configured_extension(
            Path::new("Makefile"),
            &extensions
        ));
        assert!(!matches_configured_extension(
            Path::new("notes.txt.bak"),
            &extensions
        ));
    }

    /// A dot file has no extension — the dot begins its name — so the extension list
    /// is read against the name without that dot. This is what makes `gitignore` in
    /// the list mean `.gitignore`.
    #[test]
    fn a_dot_file_is_matched_by_the_name_it_is_written_with() {
        let extensions = sanitize_extensions("gitignore,gitattributes");
        assert!(matches_configured_extension(
            Path::new(".gitignore"),
            &extensions
        ));
        assert!(matches_configured_extension(
            Path::new(".gitattributes"),
            &extensions
        ));
        // A dot file with an extension of its own is still read by that extension.
        assert!(matches_configured_extension(
            Path::new(".eslintrc.json"),
            &sanitize_extensions("json")
        ));
        assert!(!matches_configured_extension(
            Path::new(".eslintrc.json"),
            &extensions
        ));
        // And a name that only looks like one is not a match.
        assert!(!matches_configured_extension(
            Path::new("gitignore"),
            &extensions
        ));
    }

    /// The second list is for the files that have no extension to match at all.
    #[test]
    fn names_cover_the_files_a_repository_is_recognized_by() {
        let names = sanitize_names("license,makefile,.gitignore");

        assert!(matches_configured_name(Path::new("LICENSE"), &names));
        assert!(matches_configured_name(Path::new("License"), &names));
        assert!(matches_configured_name(Path::new("Makefile"), &names));
        assert!(matches_configured_name(
            Path::new("C:\\projects\\thing\\.gitignore"),
            &names
        ));
        assert!(!matches_configured_name(Path::new("LICENSE.md"), &names));
        assert!(!matches_configured_name(Path::new("notes.txt"), &names));

        // An extension and a name are separate halves of the same gate, so the file
        // the user is on is matched by whichever list carries it.
        let extensions = sanitize_extensions("md");
        assert!(!matches_configured_extension(
            Path::new("LICENSE"),
            &extensions
        ));
        assert!(matches_configured_extension(
            Path::new("LICENSE.md"),
            &extensions
        ));
    }

    /// The built-in name list is written the way the lookups read it: lowercase,
    /// without the leading dot, and with nothing in it that could not be a file
    /// name.
    #[test]
    fn default_names_are_normalized_and_match_the_files_they_are_for() {
        let names = sanitize_names(DEFAULT_TEXT_NAMES);
        assert!(!names.is_empty());

        let mut sorted = names.clone();
        sorted.sort_unstable();
        sorted.dedup();
        assert_eq!(
            sorted.len(),
            names.len(),
            "the default list repeats an entry"
        );
        assert!(names.iter().all(|name| *name == name.to_lowercase()));
        assert!(names.iter().all(|name| !name.starts_with('.')));

        for file in [
            "LICENSE",
            "LICENCE",
            "Makefile",
            "Dockerfile",
            ".gitignore",
            ".gitattributes",
            ".dockerignore",
            ".editorconfig",
            ".mailmap",
            "README",
            "CHANGELOG",
            "CONTRIBUTING",
            "CMakeLists.txt",
        ] {
            assert!(
                matches_configured_name(Path::new(file), &names),
                "{file} should be previewed by its name"
            );
        }
    }

    /// The two lists are merged by the gate itself: `is_text_file` is true for a
    /// file either list claims, and the toggle turns both of them off together.
    #[test]
    fn the_gate_takes_either_list() {
        let extensions = sanitize_extensions("txt");
        let names = sanitize_names("license,.gitignore");
        let claims = |file: &str| {
            matches_configured_extension(Path::new(file), &extensions)
                || matches_configured_name(Path::new(file), &names)
        };

        assert!(claims("notes.txt"));
        assert!(claims("LICENSE"));
        assert!(claims(".gitignore"));
        assert!(!claims("archive.zip"));
        assert!(!claims("photo.png"));
    }
}
