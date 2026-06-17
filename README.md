# anno2vcf

Convert ANNO Excel workbooks to VCF from the command line.

This tool is designed for Excel annotation files with columns such as:

```text
#CHROM  POS  REF  ALT  DP  AD  QUAL  MQ  Zygosity  FILTER  Effect  Putative_Impact ...
```

It has no third-party Python dependencies. It reads `.xlsx` files directly using
Python's standard library.

## Windows CMD: Download ZIP, No Git Required

1. Download this repository as a ZIP file from GitHub.
2. Extract the ZIP.
3. Open `cmd` inside the extracted folder.
4. Run:

```bat
anno2vcf.bat examples\simple_ANNO.xlsx -o simple.vcf
```

This mode does not install anything. It only requires Python 3.9 or newer.

To install the command globally from the extracted folder:

```bat
py -m pip install .
```

Then the command becomes available as:

```bat
anno2vcf input_ANNO.xlsx -o output.vcf
```

If `anno2vcf` is not recognized after installation, close and reopen `cmd`.

## Windows CMD: Install From GitHub

Install Python 3.9 or newer first. On Windows, select **Add python.exe to PATH**
during installation.

If Git is installed, install directly from GitHub:

```bat
py -m pip install git+https://github.com/hxtunq/anno2vcf.git
```

After installation, use it from `cmd`:

```bat
anno2vcf input_ANNO.xlsx -o output.vcf
```

If `anno2vcf` is not recognized, try:

```bat
py -m anno_excel_to_vcf input_ANNO.xlsx -o output.vcf
```

Upgrade later with:

```bat
py -m pip install --upgrade --force-reinstall git+https://github.com/hxtunq/anno2vcf.git
```

Uninstall:

```bat
py -m pip uninstall anno2vcf
```

## Install on Linux or macOS

```bash
python3 -m pip install git+https://github.com/hxtunq/anno2vcf.git
```

Then run:

```bash
anno2vcf input_ANNO.xlsx -o output.vcf
```

## Use Without Installing

Clone or download this repository, then run:

```bash
python3 anno2vcf.py examples/simple_ANNO.xlsx -o simple.vcf
```

On Windows:

```bat
py anno2vcf.py examples\simple_ANNO.xlsx -o simple.vcf
```

The repository includes `examples/simple_ANNO.xlsx`, a small synthetic workbook
with one `SNP_Indel_ANNO` sheet. Use it for testing only.

## Common Commands

List sheets:

```bash
anno2vcf input_ANNO.xlsx --list-sheets
```

Convert a specific sheet:

```bash
anno2vcf input_ANNO.xlsx --sheet SNP_Indel_ANNO -o output.vcf
```

Set sample name:

```bash
anno2vcf input_ANNO.xlsx --sample SAMPLE1 -o output.vcf
```

Add reference contig lengths:

```bash
anno2vcf input_ANNO.xlsx --reference hg38 -o output.vcf
```

Available reference modes:

```text
generic       chr1..chr22, chrX, chrY, chrM without lengths
hg19/GRCh37   UCSC hg19 contigs with lengths
hg38/GRCh38   UCSC hg38 contigs with lengths
none          no contig headers
```

Default is `generic`, which avoids guessing the wrong genome build.

## Output Mapping

```text
#CHROM, POS, REF, ALT      -> VCF core columns
dbSNP142_ID                -> VCF ID
QUAL, FILTER               -> VCF QUAL and FILTER
DP, AD, MQ, annotations    -> VCF INFO
Zygosity                   -> VCF genotype GT
```

`Zygosity` is converted as:

```text
HET  -> 0/1
HOM  -> 1/1
HEMI -> 1
```

The Excel `AD` column is treated as alternate allele depth. The VCF sample `AD`
field is computed as `ref_depth,alt_depth` using `DP - AD`.

## Filter Examples With bcftools

Get variants with moderate impact:

```bash
bcftools view -i 'INFO/IMPACT="MODERATE"' simple.vcf > moderate.vcf
```

Get variants with moderate or high impact:

```bash
bcftools view -i 'INFO/IMPACT="MODERATE" || INFO/IMPACT="HIGH"' simple.vcf > moderate_high.vcf
```

Get a specific genomic position:

```bash
bcftools view -i 'CHROM="chr1" && POS=12345' simple.vcf
```

## Privacy Note

Do not commit real patient Excel or VCF files to a public GitHub repository.
Use small synthetic examples for testing and documentation.
