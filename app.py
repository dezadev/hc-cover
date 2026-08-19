import os
import sys
import re
import argparse
import importlib.util
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Tuple, Optional

HAS_PYMUPDF = importlib.util.find_spec("fitz") is not None
if HAS_PYMUPDF:
    import fitz  # PyMuPDF for high-res PDF rendering and text extraction

HAS_WIN32COM = (
    importlib.util.find_spec("win32com") is not None
    and importlib.util.find_spec("win32com.client") is not None
)
if HAS_WIN32COM:
    import win32com.client  # Windows COM Automation for CorelDRAW X7


class AutoCoverEngine:
    """Core Engine for AutoCover Studio with PDF text parsing & CorelDRAW Automation."""
    
    PRESETS: Dict[str, Tuple[float, float]] = {
        "A4": (210.0, 297.0),
        "B5": (182.0, 257.0),
        "A5": (148.0, 210.0)
    }

    PAGE_A3_PLUS_W: float = 297.0
    PAGE_A3_PLUS_H: float = 430.0

    def __init__(self, cover_size_type: str = "B5", custom_w: float = 182.0, custom_h: float = 257.0,
                 spine_width: float = 12.0, gap_width: float = 5.0, logo_path: str = "", save_individual_cdr: bool = True):
        self.cover_size_type = cover_size_type
        self.spine_width = spine_width
        self.gap_width = gap_width
        self.logo_path = logo_path
        self.save_individual_cdr = save_individual_cdr

        if cover_size_type in self.PRESETS:
            self.cover_w, self.cover_h = self.PRESETS[cover_size_type]
        else:
            self.cover_w = custom_w
            self.cover_h = custom_h

    def calculate_layout(self) -> Dict[str, float]:
        """Calculates centered layout bounds on A3+ paper (297 x 430 mm)."""
        if self.cover_w <= 0 or self.cover_h <= 0:
            raise ValueError("Ukuran cover harus lebih besar dari 0 mm.")
        if self.spine_width < 0 or self.gap_width < 0:
            raise ValueError("Tebal punggung dan gap tidak boleh bernilai negatif.")

        total_layout_w = self.spine_width + self.gap_width + self.cover_w
        if total_layout_w > self.PAGE_A3_PLUS_W:
            raise ValueError(
                f"Total lebar layout ({total_layout_w:.1f} mm) melebihi lebar A3+ "
                f"({self.PAGE_A3_PLUS_W:.1f} mm)."
            )
        if self.cover_h > self.PAGE_A3_PLUS_H:
            raise ValueError(
                f"Tinggi cover ({self.cover_h:.1f} mm) melebihi tinggi A3+ "
                f"({self.PAGE_A3_PLUS_H:.1f} mm)."
            )

        left_margin = (self.PAGE_A3_PLUS_W - total_layout_w) / 2.0
        top_margin = (self.PAGE_A3_PLUS_H - self.cover_h) / 2.0

        spine_x = left_margin
        gap_x = spine_x + self.spine_width
        cover_x = gap_x + self.gap_width
        start_y = self.PAGE_A3_PLUS_H - top_margin

        return {
            "total_w": total_layout_w,
            "left_margin": left_margin,
            "top_margin": top_margin,
            "spine_x": spine_x,
            "gap_x": gap_x,
            "cover_x": cover_x,
            "start_y": start_y
        }

    @staticmethod
    def extract_spine_components(pdf_path: str) -> Dict[str, str]:
        filename = os.path.splitext(os.path.basename(pdf_path))[0]
        normalized_filename = filename.replace("_", " ")
        clean_fn = re.sub(r'[^a-zA-Z0-9_\- ]', '', normalized_filename).strip().upper()

        # Default smart parsing from filename
        default_author = "NAMA PENULIS"
        default_title = clean_fn if clean_fn else "JUDUL SKRIPSI"

        if "-" in clean_fn:
            parts = clean_fn.split("-")
            default_author = parts[0].strip()
            default_title = " ".join(parts[1:]).strip()

        data = {
            "kategori": "SKRIPSI",
            "judul": default_title,
            "identitas": f"{default_author}\nNIM: 00000000",
            "tahun": "2026",
            "nama_file_penulis": default_author.replace(" ", "_"),
            "pdf_path": pdf_path
        }

        if not HAS_PYMUPDF:
            return data

        try:
            with fitz.open(pdf_path) as doc:
                if len(doc) == 0:
                    return data

                page = doc[0]
                text = page.get_text("text")
            lines = [line.strip() for line in text.split("\n") if line.strip()]

            if not lines:
                return data

            # Extract Year
            year_match = re.search(r'\b(20\d{2})\b', text)
            if year_match:
                data["tahun"] = year_match.group(1)

            # Detect Category
            if re.search(r'TESIS', text, re.I):
                data["kategori"] = "TESIS"
            elif re.search(r'DISERTASI', text, re.I):
                data["kategori"] = "DISERTASI"
            elif re.search(r'LAPORAN', text, re.I):
                data["kategori"] = "LAPORAN TUGAS AKHIR"
            elif re.search(r'PUBLIKASI ILMIAH', text, re.I):
                data["kategori"] = "PUBLIKASI ILMIAH"
            else:
                data["kategori"] = "SKRIPSI"

            # Extract Author Name & NIM
            found_penulis = []
            for i, l in enumerate(lines):
                if re.search(r'DISUSUN OLEH|OLEH\s*:', l, re.I):
                    if i + 1 < len(lines):
                        found_penulis.append(lines[i + 1])
                    if i + 2 < len(lines) and (re.search(r'\d+', lines[i + 2]) or re.search(r'NIM|NPM|PROGRAM|ILMU', lines[i + 2], re.I)):
                        found_penulis.append(lines[i + 2])
                    break
                elif re.search(r'\b(NIM|NPM|Nim)\b', l, re.I) or re.match(r'^\d{8,12}$', l):
                    if i > 0 and lines[i - 1] not in found_penulis:
                        found_penulis.insert(0, lines[i - 1])
                    found_penulis.append(l)
                    break

            if found_penulis:
                data["identitas"] = "\n".join(found_penulis)
                clean_author = re.sub(r'[^a-zA-Z0-9_\- ]', '', found_penulis[0]).strip().replace(" ", "_")
                if clean_author:
                    data["nama_file_penulis"] = clean_author

            # Extract Title
            title_candidates = []
            for l in lines:
                u = l.upper()
                if (len(l) > 3 and 
                    not any(kw in u for kw in ["SKRIPSI", "TESIS", "DISERTASI", "UNIVERSITAS", "FAKULTAS", "PROGRAM", "DISUSUN", "OLEH", "NIM", "NPM", "PUBLIKASI ILMIAH"]) and
                    not re.search(r'\b20\d{2}\b', u)):
                    title_candidates.append(l)

            if title_candidates:
                data["judul"] = " ".join(title_candidates[:4]).upper()

            return data
        except Exception as e:
            print(f"Peringatan ekstraksi teks PDF {pdf_path}: {e}")
            return data

    @staticmethod
    def render_pdf_cover_to_image(pdf_path: str) -> Optional[str]:
        """Renders page 1 of PDF to a high-res 300 DPI PNG file for lossless CorelDRAW import."""
        if not HAS_PYMUPDF:
            return None
        try:
            with fitz.open(pdf_path) as doc:
                if len(doc) == 0:
                    return None
                page = doc[0]
                pix = page.get_pixmap(dpi=300)

            temp_dir = tempfile.gettempdir()
            safe_name = re.sub(r"[^a-zA-Z0-9_.-]", "_", os.path.basename(pdf_path))
            out_img = os.path.join(temp_dir, f"cover_temp_{os.getpid()}_{safe_name}.png")
            pix.save(out_img)
            return out_img
        except Exception as err:
            print(f"Gagal render PNG dari PDF {pdf_path}: {err}")
            return None

    def build_coreldraw_document(self, pdf_files: List[str], output_dir: str = "", progress_callback=None) -> bool:
        if not HAS_WIN32COM:
            raise RuntimeError("Library 'pywin32' (win32com) tidak ditemukan. Mohon install: pip install pywin32")

        try:
            corel = win32com.client.Dispatch("CorelDRAW.Application.17")
        except Exception:
            try:
                corel = win32com.client.Dispatch("CorelDRAW.Application")
            except Exception as err:
                raise RuntimeError("Gagal terhubung dengan CorelDRAW X7.") from err

        corel.Visible = True
        doc = corel.CreateDocument()
        doc.Unit = 4  # cdrMillimeter
        doc.MasterPage.SetSize(self.PAGE_A3_PLUS_W, self.PAGE_A3_PLUS_H)

        layout = self.calculate_layout()

        for idx, pdf_path in enumerate(pdf_files, start=1):
            spine_data = self.extract_spine_components(pdf_path)
            author_file_name = spine_data["nama_file_penulis"]

            if progress_callback:
                progress_callback(idx, len(pdf_files), pdf_path, author_file_name)

            if idx > 1:
                page = doc.AddPages(1)
            else:
                page = doc.Pages(1)
            
            page.Name = f"{author_file_name[:25]}"
            layer = page.ActiveLayer

            # High Resolution Cover Image Import
            cover_img_path = self.render_pdf_cover_to_image(pdf_path) or pdf_path

            imp_opt = corel.CreateStructImportOptions()
            imp_opt.Mode = 1  # Full Import
            
            imp_filter = layer.ImportEx(cover_img_path, 0, imp_opt)
            imp_filter.Finish()

            sel = corel.ActiveSelectionRange
            if sel.Count > 0:
                cover_shape = sel.Group()
                cover_shape.SetSize(self.cover_w, self.cover_h)
                cover_shape.SetPosition(layout["cover_x"], layout["start_y"])

            # Clean temp image file
            if cover_img_path != pdf_path and os.path.exists(cover_img_path):
                try:
                    os.remove(cover_img_path)
                except Exception:
                    pass

            # Draw Spine Box & Gap Line
            spine_rect = layer.CreateRectangle(
                layout["spine_x"], layout["start_y"], layout["spine_x"] + self.spine_width, layout["start_y"] - self.cover_h
            )
            spine_rect.Outline.Color.FromRGBEx(0, 0, 0)
            spine_rect.Outline.Width = 0.25

            gap_line = layer.CreateLineSegment(
                layout["gap_x"], layout["start_y"], layout["gap_x"], layout["start_y"] - self.cover_h
            )
            gap_line.Outline.Color.FromRGBEx(180, 180, 180)
            gap_line.Outline.Width = 0.15

            # Rotated Spine Text (Exactly 270 Degrees)
            center_x = layout["spine_x"] + (self.spine_width / 2.0)
            
            kat_text = layer.CreateArtisticText(
                center_x, layout["start_y"] - (self.cover_h * 0.15), spine_data["kategori"], 0, 0, "Arial", 10
            )
            kat_text.Text.Angle = 270.0
            kat_text.Text.Story.Font = "Arial-Bold"

            judul_str = spine_data["judul"] if spine_data["judul"] else "JUDUL SKRIPSI"
            judul_text = layer.CreateArtisticText(
                center_x, layout["start_y"] - (self.cover_h * 0.45), judul_str, 0, 0, "Arial", 9
            )
            judul_text.Text.Angle = 270.0
            judul_text.Text.Story.Font = "Arial-Bold"

            id_str = spine_data["identitas"] if spine_data["identitas"] else "NAMA PENULIS"
            id_text = layer.CreateArtisticText(
                center_x, layout["start_y"] - (self.cover_h * 0.75), id_str, 0, 0, "Arial", 8
            )
            id_text.Text.Angle = 270.0

            thn_text = layer.CreateArtisticText(
                center_x, layout["start_y"] - self.cover_h + 12.0, spine_data["tahun"], 0, 0, "Arial", 10
            )
            thn_text.Text.Angle = 270.0
            thn_text.Text.Story.Font = "Arial-Bold"

            if self.save_individual_cdr and output_dir:
                out_path = os.path.join(output_dir, f"Hardcover_{author_file_name}.cdr")
                try:
                    doc.SaveAs(out_path)
                except Exception as save_err:
                    print(f"Simpan CDR {out_path} gagal: {save_err}")

        return True


def collect_pdf_paths(input_paths: List[str]) -> List[str]:
    """Expand file/folder inputs into a sorted list of PDF paths."""
    pdf_paths: List[str] = []
    for input_path in input_paths:
        normalized = input_path.strip('"\'')
        if os.path.isdir(normalized):
            pdf_paths.extend(
                os.path.join(normalized, f)
                for f in sorted(os.listdir(normalized))
                if f.lower().endswith(".pdf")
            )
        elif os.path.isfile(normalized) and normalized.lower().endswith(".pdf"):
            pdf_paths.append(normalized)
    return pdf_paths


def run_cli_mode(pdf_paths: Optional[List[str]] = None, spine_width: Optional[float] = None, cover_preset: Optional[str] = None, automate: bool = True):
    """Run AutoCover Studio engine in headless Command Line Interface (CLI) mode."""
    print("=" * 65)
    print(" 📘 AutoCover Studio - Headless / CLI Batch Mode")
    print("=" * 65)

    if not pdf_paths:
        cli_args = [arg for arg in sys.argv[1:] if not arg.startswith("-")]
        if cli_args:
            pdf_paths = collect_pdf_paths(cli_args)
        else:
            try:
                input_path = input("\nMasukkan path file PDF atau folder PDF: ").strip('"\'')
            except (KeyboardInterrupt, EOFError):
                print("\nProses dibatalkan.")
                return
            pdf_paths = collect_pdf_paths([input_path])

    if not pdf_paths:
        print("❌ Tidak ada file PDF untuk diproses.")
        return

    print(f"\n[+] Ditemukan {len(pdf_paths)} file PDF.")
    
    if spine_width is None:
        try:
            spine_input = input("Masukkan Tebal Punggung / Spine (mm) [default: 12.0]: ").strip()
            spine_w = float(spine_input) if spine_input else 12.0
        except (KeyboardInterrupt, EOFError):
            print("\nProses dibatalkan.")
            return
    else:
        spine_w = spine_width

    if cover_preset is None:
        try:
            preset_choice = input("Pilih Preset Cover [1: A4, 2: B5 (Default), 3: A5]: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nProses dibatalkan.")
            return
        preset_map = {"1": "A4", "2": "B5", "3": "A5"}
        cover_preset = preset_map.get(preset_choice, "B5")

    engine = AutoCoverEngine(
        cover_size_type=cover_preset,
        spine_width=spine_w,
        gap_width=5.0
    )

    print("\n--- Hasil Ekstraksi Metadata PDF ---")
    for idx, pdf in enumerate(pdf_paths, start=1):
        meta = AutoCoverEngine.extract_spine_components(pdf)
        print(f"[{idx}/{len(pdf_paths)}] {os.path.basename(pdf)}")
        print(f"   ├─ Kategori : {meta['kategori']}")
        print(f"   ├─ Judul    : {meta['judul']}")
        print(f"   ├─ Penulis  : {meta['identitas'].replace(chr(10), ' | ')}")
        print(f"   └─ Tahun    : {meta['tahun']}\n")

    if automate and HAS_WIN32COM:
        try:
            print("🚀 Menjalankan Otomatisasi CorelDRAW...")
            engine.build_coreldraw_document(pdf_paths)
            print("✅ Selesai mengolah dokumen ke CorelDRAW!")
        except Exception as e:
            print(f"❌ CorelDRAW Error: {e}")
    else:
        print("ℹ️ CorelDRAW (win32com) tidak terdeteksi pada sistem ini.")
        print("✅ Parsing metadata PDF berhasil diselesaikan dalam mode CLI.")


class AutoCoverAppGUI(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("AutoCover Studio X7 - PDF Hardcover Generator")
        self.geometry("880x720")
        self.configure(bg="#0f172a")

        self.pdf_files: List[str] = []
        self.output_dir_var = tk.StringVar(value=os.path.expanduser("~"))

        self.setup_styles()
        self.build_ui()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TFrame", background="#1e293b")
        self.style.configure("TLabel", background="#1e293b", foreground="#f8fafc", font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"), foreground="#6366f1")

    def build_ui(self):
        header_frame = tk.Frame(self, bg="#0f172a", pad=12)
        header_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            header_frame,
            text="📘 AutoCover Studio X7 (Rotasi 270° & PDF Embed)",
            font=("Segoe UI", 14, "bold"),
            fg="#6366f1",
            bg="#0f172a"
        )
        title_lbl.pack(anchor="w")

        main_frame = ttk.Frame(self, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        left_col = ttk.Frame(main_frame)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        ttk.Label(left_col, text="1. Ukuran Cover Depan", style="Header.TLabel").pack(anchor="w", pady=(0, 4))
        self.preset_var = tk.StringVar(value="B5")
        preset_combo = ttk.Combobox(
            left_col, textvariable=self.preset_var,
            values=["A4 (210 x 297 mm)", "B5 (182 x 257 mm)", "A5 (148 x 210 mm)"],
            state="readonly", width=28
        )
        preset_combo.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(left_col, text="2. Punggung & Gap", style="Header.TLabel").pack(anchor="w", pady=(10, 4))
        spine_frame = ttk.Frame(left_col)
        spine_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(spine_frame, text="Tebal Punggung (mm):").pack(side=tk.LEFT)
        self.ent_spine = ttk.Entry(spine_frame, width=8)
        self.ent_spine.insert(0, "12.0")
        self.ent_spine.pack(side=tk.RIGHT)

        right_col = ttk.Frame(main_frame)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(right_col, text="Daftar File PDF Cover", style="Header.TLabel").pack(anchor="w", pady=(0, 4))
        btn_select = tk.Button(
            right_col, text="📂 Pilih Banyak File PDF...", command=self.select_pdf_files,
            bg="#4f46e5", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=6
        )
        btn_select.pack(fill=tk.X, pady=(0, 8))

        list_frame = ttk.Frame(right_col)
        list_frame.pack(fill=tk.BOTH, expand=True)
        self.pdf_listbox = tk.Listbox(list_frame, bg="#0f172a", fg="#f1f5f9", font=("Segoe UI", 9))
        self.pdf_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.btn_run = tk.Button(
            self, text="🚀 PROSES BANYAK PDF KE COREL DRAW (ROTASI 270°)",
            command=self.run_automation, bg="#10b981", fg="#ffffff", font=("Segoe UI", 11, "bold"), relief="flat", pady=10
        )
        self.btn_run.pack(fill=tk.X, padx=12, pady=(0, 12))

    def select_pdf_files(self):
        files = filedialog.askopenfilenames(title="Pilih File PDF Cover Batch", filetypes=[("PDF Documents", "*.pdf")])
        if files:
            self.pdf_files = list(files)
            self.pdf_listbox.delete(0, tk.END)
            for f in self.pdf_files:
                meta = AutoCoverEngine.extract_spine_components(f)
                penulis_head = meta["identitas"].split("\n")[0]
                self.pdf_listbox.insert(tk.END, f"📄 {penulis_head} | Judul: {meta['judul'][:35]}...")

    def run_automation(self):
        if not self.pdf_files:
            messagebox.showwarning("Peringatan", "Silakan pilih minimal 1 file PDF!")
            return

        engine = AutoCoverEngine(
            cover_size_type=self.preset_var.get().split()[0],
            spine_width=float(self.ent_spine.get()),
            gap_width=5.0
        )

        try:
            engine.build_coreldraw_document(self.pdf_files, output_dir=self.output_dir_var.get())
            messagebox.showinfo("Berhasil!", f"Selesai mengolah {len(self.pdf_files)} file PDF ke CorelDRAW X7!")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="AutoCover Studio X7 - generator layout hardcover PDF.")
    parser.add_argument("paths", nargs="*", help="File PDF atau folder berisi PDF.")
    parser.add_argument("--cli", "-c", action="store_true", help="Jalankan mode command line.")
    parser.add_argument("--spine", type=float, default=None, help="Tebal punggung buku dalam mm.")
    parser.add_argument("--preset", choices=sorted(AutoCoverEngine.PRESETS), default=None, help="Preset ukuran cover depan.")
    parser.add_argument("--no-corel", action="store_true", help="Hanya ekstrak metadata; jangan jalankan otomatisasi CorelDRAW.")
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = parse_args(sys.argv[1:])
    pdf_paths = collect_pdf_paths(args.paths) if args.paths else None

    # Cek apakah flag --cli digunakan atau jika berjalan di lingkungan tanpa layar ($DISPLAY)
    if args.cli:
        run_cli_mode(pdf_paths=pdf_paths, spine_width=args.spine, cover_preset=args.preset, automate=not args.no_corel)
    else:
        try:
            app = AutoCoverAppGUI()
            app.mainloop()
        except tk.TclError as err:
            err_str = str(err)
            if "no display name" in err_str or "DISPLAY" in err_str or "couldn't connect to display" in err_str:
                print("\n⚠️ Lingkungan tanpa Layar/GUI terdeteksi (Tidak ada $DISPLAY).")
                print("🔄 Beralih otomatis ke Mode Command Line (CLI)...\n")
                run_cli_mode(pdf_paths=pdf_paths, spine_width=args.spine, cover_preset=args.preset, automate=not args.no_corel)
            else:
                raise err