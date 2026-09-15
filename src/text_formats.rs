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

/// Whether `path` carries an extension the configuration previews as text.
pub fn matches_configured_extension(path: &Path, extensions: &[String]) -> bool {
    let extension = match path.extension().and_then(|ext| ext.to_str()) {
        Some(extension) => extension.to_lowercase(),
        None => return false,
    };

    extensions.contains(&extension)
}

/// Whether the file is previewed as text under the current configuration. The
/// `Enable TXT Preview` toggle is checked first, so turning text previews off
/// leaves the extension list alone and turning them back on restores it.
pub fn is_text_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| {
            config.text_preview_enabled
                && matches_configured_extension(path, &config.text_extensions)
        })
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::{matches_configured_extension, sanitize_extensions, DEFAULT_TEXT_EXTENSIONS};
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
}
