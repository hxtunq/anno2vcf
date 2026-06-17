# anno2vcf

Convert ANNOVAR multianno Excel workbooks to Variant Call Format file (VCF) from the command line.

(useful for those moments when the original VCF mysteriously disappears)

This tool is designed for Excel VCF's annotation files with columns such as:

```text
#CHROM  POS  REF  ALT  DP  AD  QUAL  MQ  Zygosity  FILTER  Effect  Putative_Impact ...
```

It has no third-party Python dependencies. It reads `.xlsx` files directly using Python's standard library.

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

## Validate the output VCF

Use `PowerShell` in `cmd` to quickly inspect the converted VCF file and preview the first 10 non-header variant records.

```bash
powershell -c "gc output.vcf -TotalCount 100"
```

```bash
powershell -c "gc output.vcf | ? {$_ -notmatch '^#'} | select -First 10"
```
