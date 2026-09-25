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
pub const DEFAULT_TEXT_EXTENSIONS: &str =
    "adb,adoc,ads,asciidoc,asm,asp,aspx,astro,awk,bash,bat,bib,bzl,c,cc,cfg,cg,cjs,clj,cljc,cljs,\
cmake,cmd,comp,conf,cpp,cs,csh,cshtml,css,csv,csx,cts,cxx,d,dart,diff,diz,edn,ejs,el,elm,env,erb,\
erl,ex,exs,f,f03,f77,f90,f95,fish,for,frag,fs,fsi,fsx,ftn,fx,geom,glsl,go,gql,gradle,graphql,\
groovy,h,haml,hbs,hcl,hh,hlsl,hpp,hrl,hs,htm,html,hxx,inc,ini,ipynb,java,jl,js,json,json5,jsonc,\
jsonl,jsp,jsx,ksh,kt,kts,latex,less,lhs,liquid,lisp,ll,lock,log,lsp,lua,m,mak,man,markdown,md,\
mdown,metal,mjs,mk,mkd,ml,mli,mm,mts,mustache,nasm,nfo,nim,ninja,nix,njk,org,pas,patch,php,phtml,\
pl,plist,pm,properties,proto,ps1,psd1,psm1,py,pyi,pyw,r,rake,rb,rkt,rmd,rs,rst,rtf,s,sass,scala,\
scm,scss,sh,slim,sol,sql,srt,ss,styl,sv,svelte,svh,swift,tcl,tex,text,tf,tfvars,toml,ts,tsv,tsx,\
twig,txt,v,vbs,vert,vhd,vhdl,vtt,vue,wat,wgsl,xhtml,xml,xsd,xsl,xslt,yaml,yml,zig,zsh";

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
///
/// The entries are ordered by the name each one matches, a leading dot aside, since
/// a dot begins a name rather than changing it: `.gitattributes` sits where
/// `gitattributes` would, which is also the form this list is written to `config.ini`
/// in and looked up by.
pub const DEFAULT_TEXT_NAMES: &str = "authors,.babelrc,brewfile,caddyfile,changelog,changes,.clang-format,.clang-tidy,cmakelists.txt,\
code_of_conduct,containerfile,contributing,contributors,copying,copyright,dockerfile,\
.dockerignore,.editorconfig,.env,.env.example,.env.local,.eslintignore,.eslintrc,gemfile,\
.gitattributes,.gitconfig,.gitignore,.gitkeep,.gitmodules,gnumakefile,.golangci.yml,history,\
.htaccess,install,jenkinsfile,justfile,licence,license,.mailmap,makefile,makefile.am,makefile.in,\
notice,.npmignore,.prettierignore,.prettierrc,procfile,rakefile,readme,.rustfmt.toml,security,\
.stylelintrc,unlicense,vagrantfile";

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
///
/// It is shared rather than copied for the one other caller that has to agree with a list about
/// what a file is called: the tool a name is routed to inside an installed PeaZip (see
/// `peazip_formats::Backend::of`), which has to read a name the same way the list that claims it
/// did.
pub(crate) fn lookup_extension(path: &Path) -> Option<String> {
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

/// Whether either configured list claims `path`.
///
/// This is the classification without the gate: the lists as they stand, so a
/// caller that already holds the configuration can ask what kind of preview a
/// file is without asking whether that kind is switched on.
pub fn matches_text_lists(path: &Path, extensions: &[String], names: &[String]) -> bool {
    matches_configured_extension(path, extensions) || matches_configured_name(path, names)
}

/// Whether the file is previewed as text under the current configuration. The
/// `Text` gate is checked first, so turning text previews off leaves the lists
/// alone and turning them back on restores it.
///
/// It is the configured lists' own answer rather than the app's: a name the text lists hold
/// and an earlier list claims as well is that earlier kind, and what kind a file is, is asked
/// of the one order every side asks it in (see `routing::kind_of`). Nothing in the app is left
/// holding that question — the hook, the loader and the layout all ask the router — so no
/// binary reads this, and it is kept for the tests that contrast a name's list with the kind
/// the app gives it.
#[allow(dead_code)]
pub fn is_text_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| {
            config.text_preview_enabled
                && matches_text_lists(path, &config.text_extensions, &config.text_names)
        })
        .unwrap_or(false)
}
