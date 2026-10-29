extern crate docx_rust;

use docx_rust::{
    document::{BodyContent, ParagraphContent, RunContent},
    rels::TargetMode,
    DocxFile,
};
use std::collections::HashMap;
use std::fs::read_dir;

#[test]
fn read_and_replace() {
    // reader
    let path = std::path::Path::new("./tests/aaa/aa.docx");
    let book = DocxFile::from_file(path).unwrap();
    let mut docx = book.parse().unwrap();
    docx.document.body.replace_text_simple("好日子", "好天气");

    let path2 = std::path::Path::new("./tests/bbb/aa.docx");
    let _d = docx.write_file(path2).unwrap();
}

#[test]
fn replace_text_multiple() {
    // reader
    let path = std::path::Path::new("./tests/aaa/aa.docx");
    let book = DocxFile::from_file(path).unwrap();
    let mut docx = book.parse().unwrap();

    let map = HashMap::from([("好日子", "好天气")]);
    docx.document.body.replace_text(&map).unwrap();

    let map = HashMap::from([("好日子".to_string(), "好天气".to_string())]);
    docx.document.body.replace_text(&map).unwrap();

    let slice = [("好日子", "好天气")];
    docx.document.body.replace_text(&slice).unwrap();

    let slice = [("好日子".to_string(), "好天气".to_string())];
    docx.document.body.replace_text(&slice).unwrap();

    let vec = vec![("好日子", "好天气")];
    docx.document.body.replace_text(&vec).unwrap();

    let vec = vec![("好日子".to_string(), "好天气".to_string())];
    docx.document.body.replace_text(&vec).unwrap();
}

#[test]
fn read_list() {
    let path = std::path::Path::new("./tests/aaa/aa_list.docx");
    let book = DocxFile::from_file(path).unwrap();
    let docx = book.parse().unwrap();
    let text = docx.document.body.text();
    assert_eq!(
        "Dock。 List \r\nTest list\r\nNano editor\r\nTest\r\nNano",
        text
    );
}

#[test]
fn read_external_links() {
    let path = std::path::Path::new("./tests/pandoc/links.docx");
    let book = DocxFile::from_file(path).unwrap();
    let docx = book.parse().unwrap();
    if let Some(document_relationships) = docx.document_rels {
        assert_eq!(document_relationships.relationships.len(), 10);
        assert_eq!(
            document_relationships
                .relationships
                .iter()
                .filter(|r| {
                    match &r.target_mode {
                        Some(target_mode) => *target_mode == TargetMode::External,
                        None => false,
                    }
                })
                .collect::<Vec<_>>()
                .len(),
            2
        );
        assert_eq!(
            document_relationships.relationships.last().unwrap().target,
            "http://pandoc.org/README.html#synopsis"
        );
    } else {
        assert!(false);
    }
}

#[test]
fn read_pandocs() {
    if let Ok(dir) = read_dir("./tests/pandoc/") {
        for entry in dir {
            if let Ok(entry) = entry {
                let path = entry.path();
                // Check if the entry is a file
                if path.is_file() {
                    match DocxFile::from_file(path) {
                        Ok(docx_file) => {
                            // Process the DocxFile as needed
                            match docx_file.parse() {
                                Ok(_) => assert!(true),
                                Err(err) => assert!(false, "Error processing file: {:?}", err),
                            }
                        }
                        Err(err) => {
                            // Handle the error if DocxFile::from_file() fails
                            assert!(false, "Error processing file: {:?}", err);
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn read_image() {
    let path = std::path::Path::new("./tests/pandoc/image.docx");
    let book = DocxFile::from_file(path).unwrap();
    let docx = book.parse().unwrap();
    let mut is_first = true;
    for content in docx.document.body.content {
        if !is_first {
            return ();
        }
        match content {
            BodyContent::Paragraph(paragraph) => {
                for para_content in paragraph.content {
                    match para_content {
                        ParagraphContent::Run(run) => {
                            for run_content in run.content {
                                match run_content {
                                    RunContent::Drawing(drawing) => {
                                        is_first = false;
                                        if let Some(inline) = drawing.inline {
                                            if let Some(extent) = inline.extent {
                                                assert_eq!(1905000, extent.cx);
                                                assert_eq!(1905000, extent.cy);
                                            }

                                            if let Some(graphic) = inline.graphic {
                                                if let Some(cnvpr) =
                                                    &graphic.data.children[0].nv_pic_pr.c_nv_pr
                                                {
                                                    assert_eq!(
                                                        "lalune.jpg",
                                                        cnvpr.clone().descr.unwrap()
                                                    );
                                                    assert_eq!(22, cnvpr.id.unwrap());
                                                }
                                                assert_eq!(
                                                    "rId20",
                                                    graphic.data.children[0].fill.blip.embed
                                                );
                                                if let Some(relationships) = &docx.document_rels {
                                                    if let Some(target) =
                                                        relationships.get_target("rId20")
                                                    {
                                                        assert_eq!("media/rId20.jpg", target);
                                                    } else {
                                                        assert!(false)
                                                    }
                                                }
                                            } else {
                                                assert!(false)
                                            }
                                        }
                                        ()
                                    }
                                    _ => (),
                                }
                            }
                            ()
                        }
                        _ => (),
                    }
                }
                ()
            }
            _ => (),
        }
    }
}

#[tokio::test]
#[cfg(feature = "async")]
async fn read_write_async() {
    use tokio::fs::File;
    use tokio::io::BufReader;
    use tokio_util::compat::{TokioAsyncReadCompatExt, TokioAsyncWriteCompatExt};

    let file = File::open("./tests/aaa/aa.docx").await.unwrap();
    let reader = BufReader::new(file);
    let book = DocxFile::from_async_reader(reader.compat()).await.unwrap();
    let mut docx = book.parse().unwrap();

    let writer = Vec::new();
    let _ = docx.write_async(writer.compat_write()).await.unwrap();
}

#[test]
fn rejects_conflicting_zip_header_names() {
    use std::io::{Cursor, Read};
    let mut document = docx_rust::Docx::default();
    let bytes = document
        .write(Cursor::new(Vec::new()))
        .unwrap()
        .into_inner();
    assert!(DocxFile::from_reader(Cursor::new(bytes.clone())).is_ok());
    let mut cloaked = bytes;
    let name = "word/document.xml";
    let offset = zip::ZipArchive::new(Cursor::new(&cloaked))
        .unwrap()
        .by_name(name)
        .unwrap()
        .header_start() as usize
        + 30;
    assert_eq!(&cloaked[offset..offset + name.len()], name.as_bytes());
    cloaked[offset] = b'x';
    let mut ordinary = zip::ZipArchive::new(Cursor::new(&cloaked)).unwrap();
    let mut content = String::new();
    ordinary
        .by_name(name)
        .unwrap()
        .read_to_string(&mut content)
        .unwrap();
    assert!(!content.is_empty());
    let error = DocxFile::from_reader(Cursor::new(cloaked))
        .err()
        .expect("conflicting names must be rejected");
    assert!(error.to_string().contains("filename mismatch"), "{error}");
}

#[cfg(feature = "async")]
#[tokio::test]
async fn async_reader_rejects_conflicting_zip_header_names() {
    use std::io::Cursor;
    let mut document = docx_rust::Docx::default();
    let mut bytes = document
        .write(Cursor::new(Vec::new()))
        .unwrap()
        .into_inner();
    assert!(
        DocxFile::from_async_reader(futures_lite::io::Cursor::new(bytes.clone()))
            .await
            .is_ok()
    );
    bytes[30] = b'x';
    let error = DocxFile::from_async_reader(futures_lite::io::Cursor::new(bytes))
        .await
        .err()
        .expect("conflicting names must be rejected");
    assert!(error.to_string().contains("filename mismatch"), "{error}");
}
