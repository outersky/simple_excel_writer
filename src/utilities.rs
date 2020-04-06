use std::{
    fs::{self, File},
    io::{Error, ErrorKind, Read, Result, Write},
    path::Path,
};

use zip::{self, ZipWriter};

///zip_files takes a path an zipps it's contents to an output file. It keeps the folder structure intact.
/// It uses CompressionMethod::Stored by default
pub fn zip_files(in_path: &str, output_file: &str) -> Result<()> {
    let final_file = File::create(output_file)?;
    let mut zipwriter = zip::ZipWriter::new(final_file);

    // read the files in tmp_dir and add them to the zip archive
    zipper(&mut zipwriter, Path::new(&in_path), None)?;
    zipwriter.finish()?;
    Ok(())
}

/// Zipper walks the directory recursively and adds the files and folders found to a zipwriter
fn zipper(writer: &mut ZipWriter<File>, base_path: &Path, prefix: Option<String>) -> Result<()> {
    let prefix = if let Some(prefix) = prefix {
        format!("{}", &prefix)
    } else {
        String::new()
    };

    // Will not work with Bzip2 and excel
    let options =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    // read the directory
    let dir: fs::ReadDir = fs::read_dir(base_path)?;

    let mut buffer = vec![];
    for entry in dir {
        let entry: fs::DirEntry = entry.map_err(|e| Error::from(e))?;
        // set the current entry we are looking at
        let cur_entry = entry.path();

        let name: &str = cur_entry
            .file_name()
            .ok_or(ErrorKind::InvalidData)?
            .to_str()
            .ok_or(ErrorKind::InvalidData)?;

        // add the prefix if any to the current path for the entry so we keep the relative structure intact
        // i.e. usr/dev/excel/tmp_dir/xl/file.xml => xl/file.xml
        let name = format!("{}{}", &prefix, &name);

        if cur_entry.is_file() {
            // add the file entry to the zip writer
            writer.start_file(name, options)?;
            // open the file
            let mut f = File::open(&cur_entry)?;
            // read the file into the buffer
            f.read_to_end(&mut buffer)?;
            // write the buffered content to the zip file
            writer.write_all(&buffer)?;
            // clear the buffer
            buffer.clear();
        } else if cur_entry.is_dir() {
            // add prefix to the files in this dir
            let prefix = format!("{}/", name);
            // recursively add the files in the directory to the zipwriter
            zipper(writer, &cur_entry, Some(prefix))?;
        }
    }

    Ok(())
}
