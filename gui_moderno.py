# -*- coding: utf-8 -*-
import os, sys, glob, threading, time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# ---- UI moderna
import customtkinter as ctk

# ---- tu módulo principal (SIN cambios)
import planificador as pl


# ==============================
# Helpers de normalización
# ==============================
def _norm_local(s: str) -> str:
    try:
        return pl._norm(s)
    except Exception:
        return str(s).strip().lower() if s is not None else ""


# ==============================
# Overrides que inyecta la GUI
# ==============================
def _fake_pick_list(items, prechecked=None):
    # Selecciona todo por defecto (se usa para títulos que no son "Pozos a EXCLUIR")
    return set(items if items is not None else [])

def make_overrides_for_planificador(
    zonas_seleccionadas, baterias_dict_norm, pozos_a_excluir, equipos_activos
):
    """
    Reemplaza los pickers del planificador para que use lo elegido en la GUI.
    """

    # 1) zonas
    def _pick_zonas_checkbox_override(_series):
        labels = set(zonas_seleccionadas)
        norms = {_norm_local(x) for x in zonas_seleccionadas}
        return labels, norms

    # 2) baterías (dict: zona_norm -> set(bateria_norm) o None)
    def _pick_baterias_subfilter_override(_df_norm, _zl, _zn):
        return baterias_dict_norm

    # 3) exclusión de pozos (solo para ese título)
    def _pick_list_checkbox_override(title, items, prechecked=None):
        if str(title).lower().startswith("pozos a excluir"):
            return set(pozos_a_excluir)
        return _fake_pick_list(items, prechecked)

    # 4) cantidad de equipos (no preguntes por consola)
    def _prompt_int_override(prompt, default, lo, hi):
        try:
            v = int(equipos_activos)
        except Exception:
            v = default
        v = max(lo, min(hi, v))
        return v

    pl.pick_zonas_checkbox = _pick_zonas_checkbox_override
    pl.pick_baterias_subfilter = _pick_baterias_subfilter_override
    pl.pick_list_checkbox = _pick_list_checkbox_override
    pl.prompt_int = _prompt_int_override


# ==============================
# Lógica de pre-análisis
# ==============================
def preanalisis_paths(sw_path: str, nombres_path: str):
    """
    Corre el mismo pipeline de lectura y normalización que hace el planificador
    para mostrar ZONAS / BATERÍAS / POZOS en la GUI, sin planificar todavía.
    """
    # 1) historial
    df_hist = pl.read_historial(sw_path, pl.SHEET_HIST)
    # 2) diccionario
    key2off, dict_df = pl.load_pozo_dictionary(nombres_path)
    # 3) normalización
    df_norm, alert_table, norm_table = pl.apply_pozo_normalization(df_hist, key2off, dict_df)

    # Nos quedamos solo con POZOS válidos (igual que el main)
    df_norm = df_norm[df_norm["VALIDO_POZO"] == True].copy()

    # ZONAS (labels originales) y su versión normalizada
    zonas_labels = sorted(set(df_norm["ZONA"].dropna().astype(str).str.strip()))
    # BATERÍAS solo para la zona con subfiltro
    target_label = "Las Heras CG - Canadon Escondida"
    target_norm  = _norm_local(target_label)

    bats_target = []
    if target_norm in set(df_norm["__ZONA_NORM"]):
        bats_series = df_norm.loc[df_norm["__ZONA_NORM"] == target_norm, "BATERIA"] \
                             .dropna().astype(str).str.strip()
        bats_series = bats_series[bats_series != ""]
        bats_target = sorted(set(bats_series))

    return df_norm, zonas_labels, bats_target


def compute_output_glob(input_file: str) -> str:
    folder = Path(os.path.dirname(os.path.abspath(input_file)))
    stem   = os.path.splitext(os.path.basename(input_file))[0]
    today  = time.strftime("%Y%m%d")
    return str(folder / f"{stem}_CRONOGRAMA_{today}*.xlsx")


# ==============================
# GUI
# ==============================
class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")  # base celeste/azul
        self.title("Planificador Swabbing – GUI Moderna")
        self.geometry("1200x780")
        self.minsize(1100, 720)

        # Colores suaves (celestes/grises/negros)
        self.fg = ("#98c1ff", "#98c1ff")
        self.accent = "#00b4d8"

        # Estado
        self.sw_path = tk.StringVar(value="")
        self.nombres_path = tk.StringVar(value="")
        self.coords_path = tk.StringVar(value="")
        self.equipos_var = tk.IntVar(value=pl.DEFAULTS.get("equipos_activos", 4))
        self.radius_var = tk.DoubleVar(value=float(pl.RADIUS_KM))
        self.rm3d_var = tk.DoubleVar(value=float(pl.RM3D_MIN))

        self.df_norm = None
        self.zonas_labels = []
        self.bats_target = []
        self.selected_zonas = set()
        self.selected_bats_target = set()   # baterías elegidas para la zona target
        self.exclusion_map = {}             # pozo -> checkbox
        self.log_text = None

        self._build_ui()

    # ------------- UI building -------------
    def _build_ui(self):
        # ========== TOP: Selección de archivos ==========
        top = ctk.CTkFrame(self, corner_radius=16)
        top.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(top, text="Excel de Swabbing:", width=140).grid(row=0, column=0, padx=8, pady=8, sticky="e")
        e1 = ctk.CTkEntry(top, textvariable=self.sw_path, width=600)
        e1.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkButton(top, text="Buscar...", command=self._pick_sw).grid(row=0, column=2, padx=8, pady=8)

        ctk.CTkLabel(top, text="Nombres-Pozo:", width=140).grid(row=1, column=0, padx=8, pady=8, sticky="e")
        e2 = ctk.CTkEntry(top, textvariable=self.nombres_path, width=600)
        e2.grid(row=1, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkButton(top, text="Buscar...", command=self._pick_nom).grid(row=1, column=2, padx=8, pady=8)

        ctk.CTkLabel(top, text="Coordenadas:", width=140).grid(row=2, column=0, padx=8, pady=8, sticky="e")
        e3 = ctk.CTkEntry(top, textvariable=self.coords_path, width=600)
        e3.grid(row=2, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkButton(top, text="Buscar...", command=self._pick_coord).grid(row=2, column=2, padx=8, pady=8)

        ctk.CTkButton(
            top, text="Analizar archivos", fg_color=self.accent,
            command=self._analizar
        ).grid(row=0, column=3, rowspan=3, padx=14, pady=8, sticky="ns")

        # ========== MIDDLE: Filtros ==========
        mid = ctk.CTkFrame(self, corner_radius=16)
        mid.pack(fill="both", expand=True, padx=12, pady=10)
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_columnconfigure(1, weight=1)
        mid.grid_columnconfigure(2, weight=1)

        # ZONAS
        self.zonas_frame = ctk.CTkFrame(mid)
        self.zonas_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(self.zonas_frame, text="ZONAS (elegí 1+)", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(8,4))
        self.zonas_scroll = ctk.CTkScrollableFrame(self.zonas_frame, height=280)
        self.zonas_scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # BATERÍAS (solo para zona target)
        self.bats_frame = ctk.CTkFrame(mid)
        self.bats_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(self.bats_frame, text="BATERÍAS (solo 'Las Heras CG - Canadon Escondida')",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(8,4))
        self.bats_scroll = ctk.CTkScrollableFrame(self.bats_frame, height=280)
        self.bats_scroll.pack(fill="both", expand=True, padx=8, pady=8)
        self.bats_hint = ctk.CTkLabel(self.bats_frame, text="(Mostrará opciones tras elegir ZONA)", text_color="#9aa0a6")
        self.bats_hint.pack(pady=(2,8))

        # EXCLUSIONES
        self.excl_frame = ctk.CTkFrame(mid)
        self.excl_frame.grid(row=0, column=2, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(self.excl_frame, text="Pozos a EXCLUIR", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(8,4))
        tools = ctk.CTkFrame(self.excl_frame)
        tools.pack(fill="x", padx=8)
        ctk.CTkButton(tools, text="Marcar ninguno", width=110, command=self._excl_none).pack(side="left", padx=4, pady=6)
        ctk.CTkButton(tools, text="Marcar todo", width=110, command=self._excl_all).pack(side="left", padx=4, pady=6)
        self.search_var = tk.StringVar(value="")
        sframe = ctk.CTkFrame(self.excl_frame)
        sframe.pack(fill="x", padx=8, pady=(0,4))
        ctk.CTkLabel(sframe, text="Filtro rápido:").pack(side="left", padx=6)
        ctk.CTkEntry(sframe, textvariable=self.search_var, width=180).pack(side="left", padx=6)
        ctk.CTkButton(sframe, text="Aplicar", width=90, command=self._refill_exclusions).pack(side="left", padx=6)
        self.excl_scroll = ctk.CTkScrollableFrame(self.excl_frame, height=280)
        self.excl_scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # ========== BOTTOM: Parámetros + ejecutar ==========
        bot = ctk.CTkFrame(self, corner_radius=16)
        bot.pack(fill="x", padx=12, pady=(0,12))

        ctk.CTkLabel(bot, text="Equipos:", width=70).grid(row=0, column=0, padx=8, pady=10, sticky="e")
        self.cmb_equip = ctk.CTkOptionMenu(bot, values=[str(x) for x in range(1,5)],
                                           command=lambda v: self.equipos_var.set(int(v)))
        self.cmb_equip.set(str(self.equipos_var.get()))
        self.cmb_equip.grid(row=0, column=1, padx=8, pady=10, sticky="w")

        ctk.CTkLabel(bot, text="Radio clúster (km):").grid(row=0, column=2, padx=8, pady=10, sticky="e")
        ctk.CTkEntry(bot, textvariable=self.radius_var, width=80).grid(row=0, column=3, padx=8, pady=10, sticky="w")

        ctk.CTkLabel(bot, text="r_m3_d mínimo:").grid(row=0, column=4, padx=8, pady=10, sticky="e")
        ctk.CTkEntry(bot, textvariable=self.rm3d_var, width=80).grid(row=0, column=5, padx=8, pady=10, sticky="w")

        ctk.CTkButton(
            bot, text="Generar cronograma", fg_color=self.accent,
            command=self._generar_click
        ).grid(row=0, column=6, padx=14, pady=10)

        # LOG
        logf = ctk.CTkFrame(self)
        logf.pack(fill="both", expand=False, padx=12, pady=(0,12))
        ctk.CTkLabel(logf, text="Log / Estado:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=6)
        self.log_text = ctk.CTkTextbox(logf, height=140)
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    # ------------- Acciones -------------
    def _pick_sw(self):
        p = filedialog.askopenfilename(title="Excel de Swabbing", filetypes=[("Excel","*.xlsx")])
        if p: self.sw_path.set(p)

    def _pick_nom(self):
        p = filedialog.askopenfilename(title="Nombres-Pozo", filetypes=[("Excel","*.xlsx")])
        if p: self.nombres_path.set(p)

    def _pick_coord(self):
        p = filedialog.askopenfilename(title="Coordenadas", filetypes=[("Excel","*.xlsx")])
        if p: self.coords_path.set(p)

    def _log(self, s: str):
        try:
            self.log_text.insert("end", s + "\n")
            self.log_text.see("end")
            self.update_idletasks()
        except Exception:
            print(s)

    # --- ANALIZAR ---
    def _analizar(self):
        sw = self.sw_path.get().strip()
        nm = self.nombres_path.get().strip()
        if not (sw and os.path.exists(sw)):
            messagebox.showwarning("Atención", "Seleccioná el Excel de Swabbing.")
            return
        if not (nm and os.path.exists(nm)):
            messagebox.showwarning("Atención", "Seleccioná el Excel de Nombres-Pozo.")
            return
        self._log("Analizando archivos… (puede tardar unos segundos)")
        try:
            self.df_norm, self.zonas_labels, self.bats_target = preanalisis_paths(sw, nm)
        except Exception as e:
            messagebox.showerror("Error", f"No pude analizar:\n\n{e}")
            raise

        # PINTAR ZONAS
        for w in self.zonas_scroll.winfo_children():
            w.destroy()
        self.zona_vars = []
        for z in sorted(self.zonas_labels):
            v = tk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(self.zonas_scroll, text=z, variable=v,
                                 command=self._zonas_changed)
            cb.pack(anchor="w", padx=6, pady=2)
            self.zona_vars.append((z, v))

        # limpiar bats y excl
        for w in self.bats_scroll.winfo_children():
            w.destroy()
        self.selected_zonas = set()
        self.selected_bats_target = set()
        self._clear_exclusions()

        self._log("Zonas y baterías detectadas correctamente.")
        self._log("Tip: podés subfiltrar BATERÍAS para Las Heras CG - Canadon Escondida (si la elegís).")

    def _zonas_changed(self):
        self.selected_zonas = {z for (z,v) in self.zona_vars if v.get()}
        # Baterías visibles solo si está la zona objetivo
        target = "Las Heras CG - Canadon Escondida"
        for w in self.bats_scroll.winfo_children():
            w.destroy()
        self.selected_bats_target = set()

        if target in self.selected_zonas and self.bats_target:
            self.bats_vars = []
            for b in self.bats_target:
                v = tk.BooleanVar(value=True)  # por defecto todas
                cb = ctk.CTkCheckBox(self.bats_scroll, text=b, variable=v,
                                     onvalue=True, offvalue=False,
                                     command=self._refill_exclusions)
                cb.pack(anchor="w", padx=6, pady=2)
                self.bats_vars.append((b, v))
            self.bats_hint.configure(text=f"{len(self.bats_target)} baterías disponibles.")
        else:
            self.bats_hint.configure(text="(Mostrará opciones tras elegir ZONA)")
        # Actualizar exclusiones
        self._refill_exclusions()

    def _current_bats_target_selected(self):
        if not hasattr(self, "bats_vars"):
            return set()
        return {b for (b,v) in self.bats_vars if v.get()}

    def _clear_exclusions(self):
        for w in self.excl_scroll.winfo_children():
            w.destroy()
        self.exclusion_map = {}

    def _refill_exclusions(self):
        self._clear_exclusions()
        if self.df_norm is None or not self.selected_zonas:
            return
        znorm_sel = {_norm_local(z) for z in self.selected_zonas}
        df = self.df_norm.copy()
        df = df[df["__ZONA_NORM"].isin(znorm_sel)]

        # baterías si aplica
        target_norm = _norm_local("Las Heras CG - Canadon Escondida")
        if target_norm in znorm_sel and self.bats_target:
            bats_sel = self._current_bats_target_selected()
            bats_sel = {b for b in bats_sel} if bats_sel else set(self.bats_target)
            self.selected_bats_target = bats_sel
            if bats_sel:
                df = df[(df["__ZONA_NORM"] != target_norm) | (df["BATERIA"].isin(bats_sel))]

        # filtro por texto
        q = self.search_var.get().strip().lower()
        if q:
            df = df[df["POZO"].astype(str).str.lower().str.contains(q)]

        pozos = sorted(set(df["POZO"].dropna().astype(str).str.strip()))
        for p in pozos:
            v = tk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(self.excl_scroll, text=p, variable=v)
            cb.pack(anchor="w", padx=6, pady=1)
            self.exclusion_map[p] = v

    def _excl_none(self):
        for v in self.exclusion_map.values():
            v.set(False)

    def _excl_all(self):
        for v in self.exclusion_map.values():
            v.set(True)

    # --- GENERAR ---
    def _generar_click(self):
        sw = self.sw_path.get().strip()
        nm = self.nombres_path.get().strip()
        co = self.coords_path.get().strip()
        if not (sw and os.path.exists(sw) and nm and os.path.exists(nm) and co and os.path.exists(co)):
            messagebox.showwarning("Atención", "Completá las 3 rutas (Swabbing, Nombres-Pozo, Coordenadas).")
            return
        if not self.selected_zonas:
            messagebox.showwarning("Atención", "Elegí al menos una ZONA.")
            return

        # Compilar inputs del usuario
        zonas_sel = sorted(self.selected_zonas)
        target_norm = _norm_local("Las Heras CG - Canadon Escondida")
        bats_dict_norm = {}
        for z in zonas_sel:
            zn = _norm_local(z)
            if zn == target_norm and self.selected_bats_target:
                bats_dict_norm[zn] = {_norm_local(b) for b in self.selected_bats_target}
            else:
                bats_dict_norm[zn] = None

        pozos_excluir = [p for p, v in self.exclusion_map.items() if v.get()]
        equipos = int(self.equipos_var.get())
        radius = float(self.radius_var.get())
        rm3d   = float(self.rm3d_var.get())

        # Guardar y correr en thread
        t = threading.Thread(
            target=self._run_pipeline_safe,
            args=(sw, nm, co, zonas_sel, bats_dict_norm, pozos_excluir, equipos, radius, rm3d),
            daemon=True
        )
        t.start()

    def _run_pipeline_safe(self, sw, nm, co, zonas_sel, bats_dict_norm, pozos_excluir, equipos, radius, rm3d):
        try:
            self._run_pipeline(sw, nm, co, zonas_sel, bats_dict_norm, pozos_excluir, equipos, radius, rm3d)
        except SystemExit as e:
            messagebox.showwarning("Atención", str(e) if str(e) else "Proceso cancelado.")
        except Exception as e:
            messagebox.showerror("Error al generar", f"{e}")
            raise

    def _run_pipeline(self, sw, nm, co, zonas_sel, bats_dict_norm, pozos_excluir, equipos, radius, rm3d):
        # Inyectar rutas y parámetros al módulo principal
        pl.INPUT_FILE = sw
        pl.NOMBRES_POZO_FILE = nm
        pl.COORDS_FILE = co
        pl.RADIUS_KM = radius
        pl.RM3D_MIN = rm3d
        pl.DEFAULTS["equipos_activos"] = equipos

        self._log("=== Ejecutando planificador ===")
        self._log(f"Swabbing: {sw}")
        self._log(f"Nombres:  {nm}")
        self._log(f"Coords:    {co}")
        self._log(f"Radio: {radius} km | r_m3_d mín: {rm3d}")
        self._log(f"Zonas: {', '.join(zonas_sel)}")

        # Preparar overrides
        make_overrides_for_planificador(
            zonas_sel, bats_dict_norm, pozos_excluir, equipos
        )

        # Detectar salida antes/después
        pattern = compute_output_glob(sw)
        before = set(glob.glob(pattern))

        # Ejecutar
        pl.main()

        after = set(glob.glob(pattern))
        new_files = list(after - before)
        out_path = None
        if new_files:
            new_files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            out_path = new_files[0]

        # Mensaje final
        if out_path and os.path.exists(out_path):
            self._log(f"Listo ✅  Cronograma: {out_path}")
            messagebox.showinfo("Listo ✅", f"Cronograma generado en:\n\n{out_path}")
        else:
            self._log("Proceso finalizado (revisá la consola para la ruta exacta).")
            messagebox.showinfo("Proceso finalizado",
                                "El cronograma fue generado. Si no ves la ruta aquí, revisá la consola.")


if __name__ == "__main__":
    app = App()
    app.mainloop()


