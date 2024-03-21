use std::{env, fs, io};
use std::path::{Path, PathBuf};

fn main() {
    replace_res_dir();
}

fn replace_res_dir(){
    let output_path = get_output_path();

    let input_path = Path::new(&env::var("CARGO_MANIFEST_DIR").unwrap()).join("Исходные данные");
    let output_path = Path::new(&output_path).join("Исходные данные");

    if output_path.exists() {
        fs::remove_dir_all(&output_path).expect("BUILD.RS ERROR 1");
    }

    let res = copy_dir_all(input_path, output_path); //std::fs::copy(input_path, output_path);

    if let Err(err) = res {
        println!("cargo:warning={:#?}",err)
    }
}

fn get_output_path() -> PathBuf {
    let manifest_dir_string = env::var("CARGO_MANIFEST_DIR").unwrap();
    let build_type = env::var("PROFILE").unwrap();
    let path = Path::new(&manifest_dir_string).join("target").join(build_type);
    return PathBuf::from(path);
}

fn copy_dir_all(src: impl AsRef<Path>, dst: impl AsRef<Path>) -> io::Result<()> {
    fs::create_dir_all(&dst)?;
    for entry in fs::read_dir(src)? {
        let entry = entry?;
        let ty = entry.file_type()?;
        if ty.is_dir() {
            copy_dir_all(entry.path(), dst.as_ref().join(entry.file_name()))?;
        } else {
            fs::copy(entry.path(), dst.as_ref().join(entry.file_name()))?;
        }
    }
    Ok(())
}