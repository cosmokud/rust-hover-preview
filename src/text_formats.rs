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
/// alone and turning them back on restores them.
pub fn is_text_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| {
            config.text_preview_enabled
                && matches_text_lists(path, &config.text_extensions, &config.text_names)
        })
        .unwrap_or(false)
}
