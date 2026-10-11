from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import pymupdf
from PIL import Image, ImageOps
from pdf2docx import Converter
from tkinterdnd2 import DND_FILES, TkinterDnD


APP_NAME = "ClaraConvert"
IMAGE_FORMATS = {"PNG": "PNG", "JPG": "JPEG", "WEBP": "WEBP"}
SUPPORTED_IMAGES = {
    ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif"
}
FORMATS_BY_EXTENSION = {
    ".pdf": ("DOCX", "PNG", "JPG"),
    ".docx": ("PDF", "PNG", "JPG"),
    **{extension: ("PDF", "PNG", "JPG", "WEBP") for extension in SUPPORTED_IMAGES},
}
COLORS = {
    "window": "#171C24",
    "surface": "#232B36",
    "surface_alt": "#2B3542",
    "text": "#F3F5F7",
    "muted": "#AAB5C2",
    "accent": "#A8D5BA",
    "accent_hover": "#8BC6A3",
    "entry": "#1D242E",
    "border": "#3B4654",
}


def find_libreoffice() -> str | None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable
    if os.name == "nt":
        candidates = (
            Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
            / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))
            / "LibreOffice" / "program" / "soffice.exe",
        )
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
    return None


def docx_to_pdf(source: Path, destination: Path) -> None:
    executable = find_libreoffice()
    if not executable:
        raise RuntimeError(
            "Para converter Word (DOCX) em PDF, instale o LibreOffice neste "
            "computador. Depois, abra o ClaraConvert novamente."
        )

    with tempfile.TemporaryDirectory(prefix="claraconvert-") as profile_dir:
        profile = Path(profile_dir).as_uri()
        result = subprocess.run(
            [
                executable,
                "--headless",
                f"-env:UserInstallation={profile}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(destination.parent),
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
    generated = destination.parent / f"{source.stem}.pdf"
    if result.returncode != 0 or not generated.is_file():
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(
            "O LibreOffice não conseguiu converter o documento."
            + (f"\n\nDetalhes: {detail}" if detail else "")
        )
    if generated != destination:
        generated.replace(destination)


def render_pdf(source: Path, output_dir: Path, image_format: str) -> list[Path]:
    document = pymupdf.open(source)
    if document.page_count == 0:
        raise ValueError("O PDF não contém páginas.")
    outputs = []
    for index, page in enumerate(document, start=1):
        pixmap = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        suffix = "jpg" if image_format == "JPEG" else image_format.lower()
        output = output_dir / f"{source.stem}_pagina_{index:03d}.{suffix}"
        save_image(image, output, image_format)
        outputs.append(output)
    document.close()
    return outputs


def save_image(image: Image.Image, destination: Path, image_format: str) -> None:
    image = ImageOps.exif_transpose(image)
    if image_format == "JPEG":
        if image.mode not in ("RGB", "L"):
            background = Image.new("RGB", image.size, "white")
            if "A" in image.getbands():
                background.paste(image, mask=image.getchannel("A"))
            else:
                background.paste(image.convert("RGB"))
            image = background
        image.save(destination, format="JPEG", quality=94, optimize=True)
    else:
        image.save(destination, format=image_format)


def convert_file(source: Path, output_dir: Path, target_format: str) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    extension = source.suffix.lower()
    outputs: list[Path]

    if extension == ".pdf" and target_format == "DOCX":
        destination = output_dir / f"{source.stem}.docx"
        converter = Converter(str(source))
        try:
            converter.convert(str(destination))
        finally:
            converter.close()
        outputs = [destination]
    elif extension == ".docx" and target_format == "PDF":
        destination = output_dir / f"{source.stem}.pdf"
        docx_to_pdf(source, destination)
        outputs = [destination]
    elif extension == ".docx" and target_format in IMAGE_FORMATS:
        with tempfile.TemporaryDirectory(prefix="claraconvert-") as temp_dir:
            pdf_path = Path(temp_dir) / f"{source.stem}.pdf"
            docx_to_pdf(source, pdf_path)
            outputs = render_pdf(pdf_path, output_dir, IMAGE_FORMATS[target_format])
    elif extension == ".pdf" and target_format in IMAGE_FORMATS:
        outputs = render_pdf(source, output_dir, IMAGE_FORMATS[target_format])
    elif extension in SUPPORTED_IMAGES and target_format == "PDF":
        destination = output_dir / f"{source.stem}.pdf"
        with Image.open(source) as image:
            frames = []
            for frame_index in range(getattr(image, "n_frames", 1)):
                image.seek(frame_index)
                frame = ImageOps.exif_transpose(image.copy()).convert("RGB")
                frames.append(frame)
            if not frames:
                raise ValueError("Não foi possível ler a imagem.")
            frames[0].save(
                destination,
                "PDF",
                save_all=True,
                append_images=frames[1:],
                resolution=150.0,
            )
        outputs = [destination]
    elif extension in SUPPORTED_IMAGES and target_format in IMAGE_FORMATS:
        image_format = IMAGE_FORMATS[target_format]
        suffix = "jpg" if image_format == "JPEG" else image_format.lower()
        destination = output_dir / f"{source.stem}.{suffix}"
        with Image.open(source) as image:
            save_image(image.copy(), destination, image_format)
        outputs = [destination]
    else:
        raise ValueError("Esta combinação de arquivo e formato não é suportada.")

    missing = [str(path) for path in outputs if not path.is_file()]
    if missing:
        raise RuntimeError("A conversão terminou sem gerar todos os arquivos esperados.")
    return outputs


class ClaraConvertApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_NAME} | Conversor de arquivos")
        self.root.geometry("760x590")
        self.root.minsize(680, 540)
        self.root.configure(bg=COLORS["window"])
        self.source_paths: list[Path] = []
        self.is_converting = False

        self._configure_style()
        self._build_interface()
        self._update_formats()

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "TCombobox",
            fieldbackground=COLORS["entry"],
            background=COLORS["surface_alt"],
            foreground=COLORS["text"],
            arrowcolor=COLORS["accent"],
            bordercolor=COLORS["border"],
            padding=9,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", COLORS["entry"])],
            foreground=[("readonly", COLORS["text"])],
            selectbackground=[("readonly", COLORS["entry"])],
            selectforeground=[("readonly", COLORS["text"])],
        )
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor=COLORS["entry"],
            background=COLORS["accent"],
            bordercolor=COLORS["entry"],
            lightcolor=COLORS["accent"],
            darkcolor=COLORS["accent"],
        )

    def _label(self, parent: tk.Widget, text: str, **kwargs: object) -> tk.Label:
        kwargs.setdefault("bg", COLORS["surface"])
        kwargs.setdefault("fg", COLORS["text"])
        kwargs.setdefault("font", ("Segoe UI", 10))
        return tk.Label(parent, text=text, **kwargs)

    def _build_interface(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["window"], padx=30, pady=25)
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg=COLORS["window"])
        header.pack(fill="x", pady=(0, 20))
        tk.Label(
            header,
            text="ClaraConvert",
            bg=COLORS["window"],
            fg=COLORS["accent"],
            font=("Segoe UI Semibold", 23),
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Conversões simples, sem complicação.",
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 0))

        card = tk.Frame(
            outer,
            bg=COLORS["surface"],
            padx=24,
            pady=22,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        card.pack(fill="both", expand=True)

        self._label(card, "1. Escolha o arquivo").pack(anchor="w")
        file_area = tk.Frame(
            card,
            bg=COLORS["surface_alt"],
            padx=8,
            pady=7,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        file_area.pack(fill="x", pady=(9, 21))
        file_row = tk.Frame(file_area, bg=COLORS["surface_alt"])
        file_row.pack(fill="x")
        self.file_name = tk.StringVar(value="Nenhum arquivo selecionado")
        self.file_entry = tk.Entry(
            file_row,
            textvariable=self.file_name,
            state="readonly",
            readonlybackground=COLORS["entry"],
            fg=COLORS["muted"],
            disabledforeground=COLORS["muted"],
            insertbackground=COLORS["text"],
            relief="flat",
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        self.file_entry.pack(side="left", fill="x", expand=True, ipady=11, padx=(0, 10))
        self.choose_button = self._button(file_row, "Selecionar arquivo(s)", self._choose_files)
        self.choose_button.pack(side="right", ipady=7, padx=1)
        self.drop_hint = self._label(
            file_area,
            "Selecione um ou mais arquivos ou arraste-os para esta área.",
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            justify="left",
        )
        self.drop_hint.pack(anchor="w", padx=3, pady=(6, 0))
        for widget in (file_area, file_row, self.file_entry, self.drop_hint):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._drop_files)

        self._label(card, "2. Formato desejado").pack(anchor="w")
        self.format_var = tk.StringVar()
        self.format_box = ttk.Combobox(
            card,
            textvariable=self.format_var,
            state="readonly",
            font=("Segoe UI", 10),
        )
        self.format_box.pack(fill="x", pady=(9, 7), ipady=2)
        self.format_note = self._label(
            card,
            "Selecione primeiro um arquivo para ver os formatos disponíveis.",
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            justify="left",
        )
        self.format_note.pack(anchor="w", pady=(0, 20))

        self._label(card, "3. Pasta para salvar").pack(anchor="w")
        destination_row = tk.Frame(card, bg=COLORS["surface"])
        destination_row.pack(fill="x", pady=(9, 20))
        self.destination_var = tk.StringVar(value=str(Path.home() / "Downloads"))
        self.destination_entry = tk.Entry(
            destination_row,
            textvariable=self.destination_var,
            state="readonly",
            readonlybackground=COLORS["entry"],
            fg=COLORS["muted"],
            disabledforeground=COLORS["muted"],
            relief="flat",
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        self.destination_entry.pack(side="left", fill="x", expand=True, ipady=11, padx=(0, 10))
        self.folder_button = self._button(
            destination_row, "Escolher pasta", self._choose_folder, secondary=True
        )
        self.folder_button.pack(side="right", ipady=7, padx=1)

        bottom = tk.Frame(card, bg=COLORS["surface"])
        bottom.pack(fill="x", side="bottom")
        self.convert_button = self._button(bottom, "Converter arquivo", self._convert)
        self.convert_button.pack(side="right", ipady=9, padx=1)
        self.progress = ttk.Progressbar(
            bottom, mode="indeterminate", style="Horizontal.TProgressbar", length=160
        )
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 18), ipady=3)

        self.status_var = tk.StringVar(value="Pronto para começar")
        tk.Label(
            outer,
            textvariable=self.status_var,
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(13, 0))
        tk.Label(
            outer,
            text="Dica: a conversão de Word para PDF/imagens requer o LibreOffice instalado.",
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
            anchor="w",
        ).pack(fill="x", pady=(5, 0))

    def _button(
        self, parent: tk.Widget, text: str, command: object, secondary: bool = False
    ) -> tk.Button:
        background = COLORS["surface_alt"] if secondary else COLORS["accent"]
        foreground = COLORS["text"] if secondary else COLORS["window"]
        hover = COLORS["border"] if secondary else COLORS["accent_hover"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=background,
            fg=foreground,
            activebackground=hover,
            activeforeground=foreground,
            relief="flat",
            borderwidth=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            padx=14,
        )

    def _choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecione um ou mais arquivos para converter",
            filetypes=[
                ("Arquivos aceitos", "*.pdf *.docx *.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff *.gif"),
                ("PDF", "*.pdf"),
                ("Word (DOCX)", "*.docx"),
                ("Imagens", "*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff *.gif"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if paths:
            self._set_files([Path(path) for path in paths])

    def _drop_files(self, event: tk.Event) -> str:
        if self.is_converting:
            return "break"
        try:
            dropped_paths = [Path(path) for path in self.root.tk.splitlist(event.data)]
        except tk.TclError as error:
            messagebox.showerror(
                "Não foi possível ler os arquivos",
                f"Os arquivos arrastados não puderam ser identificados.\n\n{error}",
            )
            return "break"

        accepted = [
            path for path in dropped_paths
            if path.is_file() and path.suffix.lower() in FORMATS_BY_EXTENSION
        ]
        rejected = [path for path in dropped_paths if path not in accepted]
        combined = [*self.source_paths, *accepted]
        self._set_files(combined)
        if rejected:
            names = "\n".join(path.name for path in rejected[:5])
            if len(rejected) > 5:
                names += f"\n... e mais {len(rejected) - 5} arquivo(s)"
            messagebox.showwarning(
                "Alguns itens foram ignorados",
                "São aceitos arquivos PDF, DOCX e imagens.\n\n" + names,
            )
        return "copy"

    def _set_files(self, paths: list[Path]) -> None:
        supported = [
            path for path in paths
            if path.is_file() and path.suffix.lower() in FORMATS_BY_EXTENSION
        ]
        unique: list[Path] = []
        seen: set[str] = set()
        for path in supported:
            identity = str(path.resolve()).casefold()
            if identity not in seen:
                unique.append(path)
                seen.add(identity)
        self.source_paths = unique
        if len(unique) == 1:
            self.file_name.set(str(unique[0]))
        elif unique:
            self.file_name.set(f"{len(unique)} arquivos selecionados")
        else:
            self.file_name.set("Nenhum arquivo selecionado")
        self._update_formats()
        if unique:
            self.status_var.set(f"{len(unique)} arquivo(s) selecionado(s)")

    def _update_formats(self) -> None:
        if not self.source_paths:
            options = []
        else:
            common_formats = set(FORMATS_BY_EXTENSION[self.source_paths[0].suffix.lower()])
            for source in self.source_paths[1:]:
                common_formats.intersection_update(
                    FORMATS_BY_EXTENSION[source.suffix.lower()]
                )
            options = [
                target
                for target in FORMATS_BY_EXTENSION[self.source_paths[0].suffix.lower()]
                if target in common_formats
            ]
        self.format_box["values"] = options
        if options:
            self.format_var.set(options[0])
            if any(source.suffix.lower() == ".pdf" for source in self.source_paths):
                note = (
                    "Formatos disponíveis para todos os arquivos. Cada página de PDF "
                    "vira uma imagem; PDFs digitalizados não recebem OCR."
                )
            elif any(source.suffix.lower() == ".docx" for source in self.source_paths):
                note = (
                    "Formatos disponíveis para todos os arquivos. Word para PDF ou "
                    "imagens requer LibreOffice instalado."
                )
            else:
                note = "Formatos disponíveis para todos os arquivos selecionados."
            self.format_note.configure(text=note)
        else:
            self.format_var.set("")
            if self.source_paths:
                self.format_note.configure(
                    text="Os arquivos selecionados não têm um formato de saída em comum."
                )
            else:
                self.format_note.configure(
                    text="Selecione arquivos ou arraste-os para a área acima."
                )

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Escolha onde salvar o resultado",
            initialdir=self.destination_var.get(),
            mustexist=True,
        )
        if path:
            self.destination_var.set(path)

    def _convert(self) -> None:
        if self.is_converting:
            return
        if not self.source_paths:
            messagebox.showinfo(APP_NAME, "Selecione um arquivo antes de converter.")
            return
        target = self.format_var.get()
        if not target:
            messagebox.showinfo(APP_NAME, "Escolha o formato de saída.")
            return
        output_dir = Path(self.destination_var.get()).expanduser()
        if not output_dir.is_dir():
            messagebox.showerror(APP_NAME, "A pasta de destino não existe.")
            return

        self.is_converting = True
        self.convert_button.configure(state="disabled", text="Convertendo...")
        self.choose_button.configure(state="disabled")
        self.folder_button.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set("Convertendo arquivo. Isso pode levar alguns instantes...")

        sources = self.source_paths.copy()
        worker = threading.Thread(
            target=self._convert_worker,
            args=(sources, output_dir, target),
            daemon=True,
        )
        worker.start()

    def _convert_worker(
        self, sources: list[Path], output_dir: Path, target: str
    ) -> None:
        outputs: list[Path] = []
        failures: list[tuple[Path, str]] = []
        reserved_names: set[str] = set()
        for index, source in enumerate(sources, start=1):
            self.root.after(
                0,
                lambda current=index, name=source.name, total=len(sources): (
                    self.status_var.set(f"Convertendo {current}/{total}: {name}")
                ),
            )
            try:
                with tempfile.TemporaryDirectory(
                    prefix="claraconvert-", dir=output_dir
                ) as temp_dir:
                    temporary_outputs = convert_file(source, Path(temp_dir), target)
                    for temporary_output in temporary_outputs:
                        destination = self._unique_output_path(
                            output_dir, temporary_output.name, reserved_names
                        )
                        temporary_output.replace(destination)
                        outputs.append(destination)
            except Exception as error:
                failures.append((source, str(error)))
        self.root.after(
            0, lambda: self._conversion_finished(outputs, failures)
        )

    @staticmethod
    def _unique_output_path(
        output_dir: Path, filename: str, reserved_names: set[str]
    ) -> Path:
        original = Path(filename)
        candidate = output_dir / original.name
        number = 2
        while candidate.exists() or candidate.name.casefold() in reserved_names:
            candidate = output_dir / f"{original.stem}_{number}{original.suffix}"
            number += 1
        reserved_names.add(candidate.name.casefold())
        return candidate

    def _finish_conversion(self) -> None:
        self.is_converting = False
        self.progress.stop()
        self.convert_button.configure(state="normal", text="Converter arquivo")
        self.choose_button.configure(state="normal")
        self.folder_button.configure(state="normal")

    def _conversion_finished(
        self, outputs: list[Path], failures: list[tuple[Path, str]]
    ) -> None:
        self._finish_conversion()
        self.status_var.set(
            f"Concluído: {len(outputs)} arquivo(s) gerado(s), "
            f"{len(failures)} falha(s)"
        )
        message = "\n".join(str(path) for path in outputs[:5])
        if len(outputs) > 5:
            message += f"\n... e mais {len(outputs) - 5} arquivo(s) gerado(s)"
        if failures:
            details = "\n".join(
                f"{source.name}: {error}" for source, error in failures[:5]
            )
            if len(failures) > 5:
                details += f"\n... e mais {len(failures) - 5} falha(s)"
            message += f"\n\nFalhas:\n{details}"
        if outputs:
            messagebox.showinfo(
                "Conversão concluída",
                f"Arquivos gerados: {len(outputs)}\n\n{message}",
            )
        else:
            messagebox.showerror("Falha na conversão", message or "Nenhum arquivo foi gerado.")


def main() -> None:
    root = TkinterDnD.Tk()
    ClaraConvertApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
