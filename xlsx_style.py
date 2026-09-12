"""Add cell fills, number formats and column widths to an .xlsx file.

GDAL's XLSX driver writes data but no styling, so the spreadsheet is created
with OGR and then patched here. Uses only the standard library, which matters
because QGIS ships GDAL but no Python spreadsheet library.

Relies on GDAL writing a minimal styles.xml (a single "none" fill, cellXfs
entries that only reference number formats) and emitting cells as bare
<c r="B2"> tags with no style attribute.
"""

import re
import shutil
import zipfile

# Excel expects fill 1 to be gray125 by convention; GDAL writes only fill 0,
# so it is inserted before any fill we add.
_GRAY125 = '<fill><patternFill patternType="gray125"/></fill>'


class XlsxStyler:
    """Collects styling changes, then rewrites the archive once on save()."""

    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self._styles = z.read("xl/styles.xml").decode()
        self._sheets = {}
        self._pending_fills = []
        self._pending_xfs = []
        self._n_fills = int(re.search(r'<fills count="(\d+)"', self._styles).group(1))
        self._n_xfs = int(re.search(r'<cellXfs count="(\d+)"', self._styles).group(1))
        if self._n_fills == 1:
            self._pending_fills.append(_GRAY125)

    def add_style(self, fill=None, num_fmt=None):
        """Register a style and return the index to pass to set_cell()."""
        fill_id = 0
        if fill:
            fill_id = self._n_fills + len(self._pending_fills)
            self._pending_fills.append(
                '<fill><patternFill patternType="solid">'
                f'<fgColor rgb="FF{fill}"/><bgColor rgb="FF{fill}"/>'
                "</patternFill></fill>"
            )
        num_fmt_id = self._add_num_fmt(num_fmt) if num_fmt else 164  # 164 = GENERAL
        index = self._n_xfs + len(self._pending_xfs)
        self._pending_xfs.append(
            f'<xf numFmtId="{num_fmt_id}" fillId="{fill_id}" '
            f'applyFill="1" applyNumberFormat="1" xfId="0"/>'
        )
        return index

    def set_cell(self, sheet_index, ref, style_index):
        self._sheets.setdefault(sheet_index, {"cells": {}, "widths": None})
        self._sheets[sheet_index]["cells"][ref] = style_index

    def set_widths(self, sheet_index, widths):
        """widths: {column_number: width_in_characters}"""
        self._sheets.setdefault(sheet_index, {"cells": {}, "widths": None})
        self._sheets[sheet_index]["widths"] = widths

    def save(self):
        updates = {"xl/styles.xml": self._build_styles()}
        with zipfile.ZipFile(self.path) as z:
            for index, changes in self._sheets.items():
                member = f"xl/worksheets/sheet{index}.xml"
                updates[member] = self._patch_sheet(
                    z.read(member).decode(), changes["cells"], changes["widths"]
                )
        self._rewrite(updates)

    def _add_num_fmt(self, code):
        existing = re.search(
            rf'<numFmt formatCode="{re.escape(code)}" numFmtId="(\d+)"', self._styles
        )
        if existing:
            return int(existing.group(1))
        count = int(re.search(r'<numFmts count="(\d+)"', self._styles).group(1))
        new_id = max(int(i) for i in re.findall(r'numFmtId="(\d+)"', self._styles)) + 1
        self._styles = self._styles.replace(
            "</numFmts>",
            f'<numFmt formatCode="{code}" numFmtId="{new_id}"/></numFmts>',
        ).replace(f'<numFmts count="{count}"', f'<numFmts count="{count + 1}"')
        return new_id

    def _build_styles(self):
        styles = self._styles
        if self._pending_fills:
            total = self._n_fills + len(self._pending_fills)
            styles = styles.replace(
                "</fills>", "".join(self._pending_fills) + "</fills>"
            ).replace(f'<fills count="{self._n_fills}"', f'<fills count="{total}"')
        if self._pending_xfs:
            total = self._n_xfs + len(self._pending_xfs)
            styles = styles.replace(
                "</cellXfs>", "".join(self._pending_xfs) + "</cellXfs>"
            ).replace(f'<cellXfs count="{self._n_xfs}"', f'<cellXfs count="{total}"')
        return styles

    @staticmethod
    def _patch_sheet(xml, cells, widths):
        for ref, style_index in cells.items():
            # GDAL emits <c r="B2"> with no style attribute; the lookahead keeps
            # this idempotent if that ever changes.
            xml = re.sub(
                rf'<c r="{ref}"(?![^>]*\bs=)',
                f'<c r="{ref}" s="{style_index}"',
                xml,
            )
        if widths:
            cols = "".join(
                f'<col min="{c}" max="{c}" width="{w:.2f}" customWidth="1"/>'
                for c, w in sorted(widths.items())
            )
            if "<cols>" in xml:
                xml = re.sub(
                    r"<cols>.*?</cols>", f"<cols>{cols}</cols>", xml, flags=re.DOTALL
                )
            else:
                xml = xml.replace("<sheetData>", f"<cols>{cols}</cols><sheetData>")
        return xml

    def _rewrite(self, updates):
        # zipfile cannot edit members in place, so copy to a new archive.
        tmp = self.path + ".tmp"
        with (
            zipfile.ZipFile(self.path) as zin,
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            for item in zin.infolist():
                data = updates.get(item.filename)
                zout.writestr(
                    item, data.encode("utf-8") if data else zin.read(item.filename)
                )
        shutil.move(tmp, self.path)
