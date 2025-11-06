use std::error::Error;
use std::io;
use std::path::PathBuf;

use clap::Parser;
use rs_x2jsonl::excel_to_jsonl;

#[derive(Parser)]
#[command(version, about, long_about = None)]
struct Cli {
    /// Excel file path
    #[arg(short, long)]
    input: PathBuf,

    /// Sheet name
    #[arg(short, long)]
    sheet: String,
}

fn main() -> Result<(), Box<dyn Error>> {
    let cli = Cli::parse();

    let stdout = io::stdout();
    let writer = io::BufWriter::new(stdout.lock());

    excel_to_jsonl(&cli.input, &cli.sheet, writer)?;

    Ok(())
}
