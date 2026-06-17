#!/usr/bin/env python3
"""Convert an annotation Excel file to VCF.

This targets ANNO workbooks with columns such as:
  #CHROM, POS, REF, ALT, DP, AD, QUAL, MQ, Zygosity, FILTER, Effect...

The converter streams the .xlsx XML directly, so it does not need pandas or
openpyxl and can handle sheets that contain many formatted empty rows.
"""

from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import quote
from xml.etree import ElementTree as ET


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

MISSING = {"", ".", "-", "NA", "N/A", "#N/A", "NAN", "nan", "None"}
CORE_COLUMNS = {"#CHROM", "CHROM", "POS", "REF", "ALT", "QUAL", "FILTER"}
SPECIAL_COLUMNS = {"DP", "AD", "MQ", "ZYGOSITY", "DBSNP142_ID"}
INFO_NAME_OVERRIDES = {
    "AD": "AO",
    "ZYGOSITY": "ZYG",
    "EFFECT": "EFFECT",
    "PUTATIVE_IMPACT": "IMPACT",
    "GENE_NAME": "GENE",
    "HGVS.C": "HGVS_C",
    "HGVS.P": "HGVS_P",
    "RANK/TOTAL": "RANK_TOTAL",
}
CANONICAL_CONTIGS = [f"chr{i}" for i in range(1, 23)] + ["chrX", "chrY", "chrM"]
DEFAULT_FILTER_IDS = {
    "MG_SNP_Filter": "Filter label from ANNO Excel FILTER column",
    "MG_INDEL_Filter": "Filter label from ANNO Excel FILTER column",
}
REFERENCE_ALIASES = {
    "none": "none",
    "generic": "generic",
    "hg19": "hg19",
    "grch37": "hg19",
    "GRCh37": "hg19",
    "HG19": "hg19",
    "hg38": "hg38",
    "grch38": "hg38",
    "GRCh38": "hg38",
    "HG38": "hg38",
}
REFERENCE_CONTIG_LENGTHS = {
    "hg19": {
        "chr1": 249250621,
        "chr2": 243199373,
        "chr3": 198022430,
        "chr4": 191154276,
        "chr5": 180915260,
        "chr6": 171115067,
        "chr7": 159138663,
        "chr8": 146364022,
        "chr9": 141213431,
        "chr10": 135534747,
        "chr11": 135006516,
        "chr12": 133851895,
        "chr13": 115169878,
        "chr14": 107349540,
        "chr15": 102531392,
        "chr16": 90354753,
        "chr17": 81195210,
        "chr18": 78077248,
        "chr19": 59128983,
        "chr20": 63025520,
        "chr21": 48129895,
        "chr22": 51304566,
        "chrX": 155270560,
        "chrY": 59373566,
        "chrM": 16571,
    },
    "hg38": {
        "chr1": 248956422,
        "chr2": 242193529,
        "chr3": 198295559,
        "chr4": 190214555,
        "chr5": 181538259,
        "chr6": 170805979,
        "chr7": 159345973,
        "chr8": 145138636,
        "chr9": 138394717,
        "chr10": 133797422,
        "chr11": 135086622,
        "chr12": 133275309,
        "chr13": 114364328,
        "chr14": 107043718,
        "chr15": 101991189,
        "chr16": 90338345,
        "chr17": 83257441,
        "chr18": 80373285,
        "chr19": 58617616,
        "chr20": 64444167,
        "chr21": 46709983,
        "chr22": 50818468,
        "chrX": 156040895,
        "chrY": 57227415,
        "chrM": 16569,
    },
}


def col_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 0

    index = 0
    for char in match.group(1):
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def load_shared_strings(book: zipfile.ZipFile) -> list[str]:
    try:
        with book.open("xl/sharedStrings.xml") as handle:
            root = ET.parse(handle).getroot()
    except KeyError:
        return []

    strings: list[str] = []
    for item in root.findall(f"{NS_MAIN}si"):
        strings.append("".join(text.text or "" for text in item.iter(f"{NS_MAIN}t")))
    return strings


def list_sheets(book: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(book.read("xl/workbook.xml"))
    rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{NS_PKG_REL}Relationship")
    }

    sheets = workbook.find(f"{NS_MAIN}sheets")
    if sheets is None:
        return []

    result = []
    for sheet in sheets.findall(f"{NS_MAIN}sheet"):
        name = sheet.attrib.get("name", "")
        target = rel_targets[sheet.attrib[f"{NS_REL}id"]]
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = posixpath.normpath(posixpath.join("xl", target))
        result.append((name, path))
    return result


def sheet_path(book: zipfile.ZipFile, wanted_name: str | None) -> tuple[str, str]:
    sheets = list_sheets(book)
    if not sheets:
        raise ValueError("Workbook does not contain any sheets")

    if wanted_name is None:
        return sheets[0]

    for name, path in sheets:
        if name == wanted_name:
            return name, path

    available = ", ".join(name for name, _ in sheets)
    raise ValueError(f"Sheet '{wanted_name}' was not found. Available: {available}")


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(f"{NS_MAIN}t"))

    value = cell.find(f"{NS_MAIN}v")
    if value is None or value.text is None:
        return ""

    raw_value = value.text
    if cell_type == "s":
        return shared_strings[int(raw_value)]
    if cell_type == "b":
        return "TRUE" if raw_value == "1" else "FALSE"
    return tidy_number(raw_value)


def tidy_number(value: str) -> str:
    if not re.fullmatch(r"-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?", value):
        return value
    number = float(value)
    if number.is_integer():
        return str(int(number))
    return format(number, ".12g")


def iter_xlsx_rows(
    book: zipfile.ZipFile,
    worksheet_path: str,
    shared_strings: list[str],
):
    with book.open(worksheet_path) as handle:
        for _, elem in ET.iterparse(handle, events=("end",)):
            if elem.tag != f"{NS_MAIN}row":
                continue

            values: list[str] = []
            for cell in elem.findall(f"{NS_MAIN}c"):
                index = col_index(cell.attrib.get("r", ""))
                while len(values) <= index:
                    values.append("")
                values[index] = cell_text(cell, shared_strings).strip()

            while values and values[-1] == "":
                values.pop()
            yield values
            elem.clear()


def normalized_header_map(header: list[str]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for index, name in enumerate(header):
        clean = name.strip()
        if clean:
            mapping[clean.upper()] = index
    return mapping


def require_column(mapping: dict[str, int], *names: str) -> int:
    for name in names:
        if name.upper() in mapping:
            return mapping[name.upper()]
    raise ValueError(f"Required column not found. Tried: {', '.join(names)}")


def value_at(row: list[str], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return row[index].strip()


def is_missing(value: str) -> bool:
    return value.strip() in MISSING


def sanitize_info_id(name: str) -> str:
    key = re.sub(r"[^0-9A-Za-z_.-]+", "_", name.strip())
    key = key.strip("_.-") or "FIELD"
    if not re.match(r"[A-Za-z_]", key):
        key = f"X{key}"
    return key.upper()


def unique_info_ids(header: list[str]) -> dict[int, str]:
    used: dict[str, int] = {}
    result: dict[int, str] = {}

    for index, name in enumerate(header):
        clean = name.strip()
        upper = clean.upper()
        if not clean or upper in CORE_COLUMNS or upper == "DBSNP142_ID":
            continue

        info_id = INFO_NAME_OVERRIDES.get(upper, sanitize_info_id(clean))
        used[info_id] = used.get(info_id, 0) + 1
        if used[info_id] > 1:
            info_id = f"{info_id}_{used[info_id]}"
        result[index] = info_id

    return result


def encode_info_value(value: str) -> str:
    return quote(value.replace(" ", "_"), safe="A-Za-z0-9_.:,|/+*-")


def zygosity_to_gt(value: str) -> str:
    zyg = value.strip().upper()
    if zyg in {"HET", "HETERO", "HETEROZYGOUS"}:
        return "0/1"
    if zyg in {"HOM", "HOM_ALT", "HOMOZYGOUS", "HOMOZYGOUS_ALT"}:
        return "1/1"
    if zyg in {"HOM_REF", "HOMOZYGOUS_REF"}:
        return "0/0"
    if zyg in {"HEMI", "HEMIZYGOUS"}:
        return "1"
    return "./."


def allele_depth(dp_value: str, ao_value: str, alt_value: str) -> str:
    if is_missing(dp_value) or is_missing(ao_value):
        return "."

    alt_depths = [part.strip() for part in ao_value.split(",")]
    if not all(re.fullmatch(r"\d+", part) for part in alt_depths):
        return "."
    if not re.fullmatch(r"\d+", dp_value):
        return "."

    total_depth = int(dp_value)
    alt_depth_ints = [int(part) for part in alt_depths]
    ref_depth = total_depth - sum(alt_depth_ints)
    if ref_depth < 0:
        return "."

    if "," not in alt_value and len(alt_depth_ints) == 1:
        return f"{ref_depth},{alt_depth_ints[0]}"
    return ",".join([str(ref_depth)] + [str(depth) for depth in alt_depth_ints])


def build_info(row: list[str], info_ids: dict[int, str]) -> str:
    fields = []
    for index, info_id in info_ids.items():
        value = value_at(row, index)
        if is_missing(value):
            continue
        fields.append(f"{info_id}={encode_info_value(value)}")
    return ";".join(fields) if fields else "."


def normalize_reference(reference: str) -> str:
    normalized = REFERENCE_ALIASES.get(reference, REFERENCE_ALIASES.get(reference.lower()))
    if normalized is None:
        choices = ", ".join(sorted({"none", "generic", "hg19", "hg38", "GRCh37", "GRCh38"}))
        raise ValueError(f"Unknown reference '{reference}'. Choose one of: {choices}")
    return normalized


def write_contig_headers(out, reference: str) -> None:
    if reference == "none":
        return

    lengths = REFERENCE_CONTIG_LENGTHS.get(reference, {})
    for contig in CANONICAL_CONTIGS:
        length = lengths.get(contig)
        if length is None:
            out.write(f"##contig=<ID={contig}>\n")
        else:
            out.write(f"##contig=<ID={contig},length={length}>\n")


def sanitize_filter_id(filter_id: str) -> str:
    clean = re.sub(r"[^0-9A-Za-z_.-]+", "_", filter_id.strip())
    return clean.strip("_.-")


def write_filter_headers(out, filter_ids: list[str]) -> None:
    merged = dict(DEFAULT_FILTER_IDS)
    for filter_id in filter_ids:
        clean = sanitize_filter_id(filter_id)
        if clean and clean != "PASS":
            merged.setdefault(clean, "User-provided filter label")

    for filter_id, description in merged.items():
        out.write(f'##FILTER=<ID={filter_id},Description="{description}">\n')


def write_vcf_header(
    out,
    header: list[str],
    info_ids: dict[int, str],
    sample: str,
    reference: str,
    filter_ids: list[str],
) -> None:
    out.write("##fileformat=VCFv4.2\n")
    out.write("##source=anno2vcf\n")
    write_contig_headers(out, reference)
    write_filter_headers(out, filter_ids)
    out.write('##INFO=<ID=DP,Number=1,Type=Integer,Description="Read depth from ANNO DP column">\n')
    out.write('##INFO=<ID=AO,Number=A,Type=Integer,Description="Alternate allele depth from ANNO AD column">\n')
    out.write('##INFO=<ID=MQ,Number=1,Type=Float,Description="Mapping quality from ANNO MQ column">\n')
    out.write('##INFO=<ID=ZYG,Number=1,Type=String,Description="Zygosity from ANNO Zygosity column">\n')

    for index, info_id in info_ids.items():
        original = header[index].strip()
        if info_id in {"DP", "AO", "MQ", "ZYG"}:
            continue
        out.write(
            f'##INFO=<ID={info_id},Number=1,Type=String,'
            f'Description="Annotation column: {original}">\n'
        )

    out.write('##FORMAT=<ID=GT,Number=1,Type=String,Description="Genotype inferred from Zygosity">\n')
    out.write('##FORMAT=<ID=DP,Number=1,Type=Integer,Description="Read depth">\n')
    out.write('##FORMAT=<ID=AD,Number=R,Type=Integer,Description="Allelic depths, computed from DP and ANNO AD">\n')
    out.write("#CHROM\tPOS\tID\tREF\tALT\tQUAL\tFILTER\tINFO\tFORMAT\t" + sample + "\n")


def default_sample_name(path: Path) -> str:
    stem = path.stem
    stem = re.sub(r"_SNP_Indel_ANNO$", "", stem, flags=re.IGNORECASE)
    return stem or "SAMPLE"


def convert_anno_excel_to_vcf(
    input_xlsx: Path,
    output_vcf: Path,
    sheet: str | None,
    sample: str,
    max_empty_after_data: int,
    reference: str,
    filter_ids: list[str],
) -> int:
    reference = normalize_reference(reference)
    with zipfile.ZipFile(input_xlsx) as book, output_vcf.open("w", encoding="utf-8") as out:
        _, worksheet_path = sheet_path(book, sheet)
        shared_strings = load_shared_strings(book)

        header: list[str] | None = None
        header_map: dict[str, int] | None = None
        info_ids: dict[int, str] | None = None
        column_indexes: dict[str, int | None] = {}
        written = 0
        empty_after_data = 0

        for row in iter_xlsx_rows(book, worksheet_path, shared_strings):
            if not row:
                if written:
                    empty_after_data += 1
                    if empty_after_data >= max_empty_after_data:
                        break
                continue

            first = row[0].strip().upper()
            if first in {"#CHROM", "CHROM"}:
                if header is not None:
                    continue

                header = row
                header_map = normalized_header_map(header)
                info_ids = unique_info_ids(header)

                column_indexes = {
                    "chrom": require_column(header_map, "#CHROM", "CHROM"),
                    "pos": require_column(header_map, "POS", "POSITION"),
                    "ref": require_column(header_map, "REF"),
                    "alt": require_column(header_map, "ALT"),
                    "qual": header_map.get("QUAL"),
                    "filter": header_map.get("FILTER"),
                    "dbsnp": header_map.get("DBSNP142_ID"),
                    "dp": header_map.get("DP"),
                    "ad": header_map.get("AD"),
                    "zyg": header_map.get("ZYGOSITY"),
                }
                write_vcf_header(out, header, info_ids, sample, reference, filter_ids)
                continue

            if header is None or header_map is None or info_ids is None:
                continue

            chrom = value_at(row, column_indexes["chrom"])
            pos = value_at(row, column_indexes["pos"])
            ref = value_at(row, column_indexes["ref"])
            alt = value_at(row, column_indexes["alt"])
            if any(is_missing(value) for value in [chrom, pos, ref, alt]):
                continue

            dbsnp_id = value_at(row, column_indexes["dbsnp"])
            vcf_id = "." if is_missing(dbsnp_id) else dbsnp_id
            qual = value_at(row, column_indexes["qual"]) or "."
            filt = value_at(row, column_indexes["filter"]) or "."
            dp = value_at(row, column_indexes["dp"])
            ao = value_at(row, column_indexes["ad"])
            gt = zygosity_to_gt(value_at(row, column_indexes["zyg"]))
            ad = allele_depth(dp, ao, alt)
            info = build_info(row, info_ids)
            sample_value = f"{gt}:{dp if not is_missing(dp) else '.'}:{ad}"

            out.write(
                "\t".join([chrom, pos, vcf_id, ref, alt, qual, filt, info, "GT:DP:AD", sample_value])
                + "\n"
            )
            written += 1
            empty_after_data = 0

    return written


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="anno2vcf",
        description="Convert an ANNO .xlsx file into a VCF file.",
        epilog=(
            "Examples:\n"
            "  anno2vcf sample_ANNO.xlsx -o sample.vcf\n"
            "  anno2vcf sample_ANNO.xlsx --sheet SNP_Indel_ANNO --sample SAMPLE1 --reference hg38 -o sample.vcf\n"
            "  anno2vcf sample_ANNO.xlsx --list-sheets"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input_xlsx", type=Path)
    parser.add_argument(
        "output_vcf",
        type=Path,
        nargs="?",
        help="Output VCF path. Default: input filename with .vcf suffix.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output VCF path. Overrides the positional output_vcf argument.",
    )
    parser.add_argument(
        "--sheet",
        default="SNP_Indel_ANNO",
        help="Worksheet to convert. Use empty string to read the first sheet.",
    )
    parser.add_argument("--sample", help="Sample name for the VCF genotype column.")
    parser.add_argument(
        "--max-empty-after-data",
        type=int,
        default=1000,
        help="Stop after this many consecutive empty rows after variants start.",
    )
    parser.add_argument(
        "--reference",
        default="generic",
        help="VCF contig header mode: generic, none, hg19/GRCh37, or hg38/GRCh38. Default: generic.",
    )
    parser.add_argument(
        "--filter-id",
        action="append",
        default=[],
        help="Extra FILTER ID to declare in the VCF header. Can be used multiple times.",
    )
    parser.add_argument("--list-sheets", action="store_true", help="Print workbook sheets and exit.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)

    if args.list_sheets:
        with zipfile.ZipFile(args.input_xlsx) as book:
            for name, _ in list_sheets(book):
                print(name)
        return 0

    output_vcf = args.output or args.output_vcf or args.input_xlsx.with_suffix(".vcf")
    sheet = args.sheet if args.sheet else None
    sample = args.sample or default_sample_name(args.input_xlsx)
    written = convert_anno_excel_to_vcf(
        args.input_xlsx,
        output_vcf,
        sheet,
        sample,
        args.max_empty_after_data,
        args.reference,
        args.filter_id,
    )
    print(f"Wrote {written} variants to {output_vcf}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
