import os
import csv
import ifcopenshell
import ifcopenshell.geom
import multiprocessing
from OCC.Display.SimpleGui import init_display
from OCC.Core.AIS import AIS_Shape
from OCC.Core.Quantity import Quantity_Color, Quantity_TOC_RGB
from OCC.Core.BRepBndLib import brepbndlib
from OCC.Core.Bnd import Bnd_Box
from OCC.Core.TopExp import TopExp_Explorer
from OCC.Core.TopAbs import TopAbs_EDGE, TopAbs_FACE, TopAbs_VERTEX
from OCC.Core.BRepAdaptor import BRepAdaptor_Curve
from OCC.Core.BRepPrimAPI import BRepPrimAPI_MakePrism, BRepPrimAPI_MakeSphere
from OCC.Core.BRepAlgoAPI import BRepAlgoAPI_Cut, BRepAlgoAPI_Section
from OCC.Core.BRepBuilderAPI import BRepBuilderAPI_Transform, BRepBuilderAPI_MakeVertex, BRepBuilderAPI_MakeEdge
from OCC.Core.BRep import BRep_Tool
from OCC.Core.BRepExtrema import BRepExtrema_DistShapeShape
from OCC.Core.gp import gp_Vec, gp_Pnt, gp_Trsf

# Importe für Benutzeroberfläche
import tkinter as tk
from tkinter import filedialog
import tkinter.font as tkfont
import sys
import threading
from ctypes import windll

# Konfiguration
MIN_WALL_HEIGHT = 1.0                   # Höhe der Prüfscheibe in m
OFFSET = 0.05                           # Verschiebung des Plate rechtwinklig zur Kante um X cm
LENGTH_DIFF = OFFSET * 3                # Länge der Plate muss gekürzt werden, damit am Anfang und Ende der Plate kein Schnitt entsteht
DIST_THRESHOLD = OFFSET - 0.01          # Zur Prüfung, ob das Plate in die richtige Richtung verschoben wurde
Z_TOLERANCE = 1e-3                      # Toleranz für Z-Höhe
HEIGHT_SITE = 0.0                       # Höhe des Geländes, wird für Klassifizierung Gerüst/Geländer benötigt
GRENZE_GELAENDER = 2.0 + HEIGHT_SITE    # Höhe ab der ein Geländer erforderlich ist
GRENZE_GERUEST = 3.0 + HEIGHT_SITE      # Höhe ab der ein Gerüst erforderlich ist
RADIUS = 0.1                            # Radius der Kugeln für die Visualisierung der Koordinatenunkte (der unsicheren Kanten) welche exportiert werden


# --- Hilfsfunktionen ---

def bbox_bounds(shape):
    # Gibt xmin,ymin,zmin,xmax,ymax,zmax der Bounding-Box zurück
    # wird für die zmax Berechnung des Slabs benötigt
    bbox = Bnd_Box()
    brepbndlib.Add(shape, bbox)
    return bbox.Get()

def shape_is_empty(shape):
    # Prüft, ob nach einem Cut keine Flächen mehr vorhanden sind
    # wird für die Prüfung ob eine Plate safe oder unsafe ist benötigt
    exp = TopExp_Explorer()
    exp.Init(shape, TopAbs_FACE)
    return not exp.More()

def load_ifc_data(ifc_path, settings):
    # Lädt die IFC-Datei und extrahiert Slabs und Walls mit multiprocessing-Unterstützung
    directory = os.path.dirname(ifc_path)
    base = os.path.splitext(os.path.basename(ifc_path))[0]
    model = ifcopenshell.open(ifc_path)

    slabs = model.by_type("IfcSlab")
    slab_shapes = []
    slab_iterator = ifcopenshell.geom.iterator(settings, model, multiprocessing.cpu_count(), include=slabs)
    if slab_iterator.initialize():
        idx = 0
        print("Starte Verarbeitung der Slabs...")
        while True:
            slab_shape = slab_iterator.get()
            try:
                slab_shapes.append(slab_shape.geometry)
                print(f"Slab mit GUID {slabs[idx].GlobalId} geladen")
            except Exception as error:
                slab = slabs[idx]
                print(f"Slab mit GUID {slab.GlobalId} konnte nicht geladen werden: {error}")
            idx += 1
            if not slab_iterator.next():
                break

    walls = model.by_type("IfcWall")
    wall_shapes = []
    wall_iterator = ifcopenshell.geom.iterator(settings, model, multiprocessing.cpu_count(), include=walls)
    if wall_iterator.initialize():
        idx = 0
        print()
        print("Starte Verarbeitung der Walls...")
        while True:
            wall_shape = wall_iterator.get()
            try:
                wall_shapes.append(wall_shape.geometry)
                print(f"Wand mit GUID {walls[idx].GlobalId} geladen")
            except Exception as error:
                # Hole IfcWall Objekt mit shape.id für Fehlermeldung
                wall = walls[idx]
                print(f"Wand mit GUID {wall.GlobalId} konnte nicht geladen werden: {error}")
            idx += 1
            if not wall_iterator.next():
                break
    print()
    print("Starte Berechnung der Absturzkanten...")
    return directory, base, slab_shapes, wall_shapes


def init_visualization():
    # Initialisiert Display und Farben
    display, start_display, _, _ = init_display(size=(900, 700))
    col_red   = Quantity_Color(1.0, 0.0, 0.0, Quantity_TOC_RGB)
    col_green = Quantity_Color(0.0, 1.0, 0.0, Quantity_TOC_RGB)
    col_slab  = Quantity_Color(0.8, 0.8, 0.8, Quantity_TOC_RGB)
    col_wall  = Quantity_Color(0.5, 0.5, 0.5, Quantity_TOC_RGB)
    col_magenta = Quantity_Color(255/255, 0/255, 255/255, Quantity_TOC_RGB)  # Magenta für Eckpunkte
    return display, start_display, col_red, col_green, col_slab, col_wall, col_magenta


def get_top_edges(slab_shape, slab_z):
    # Sammelt obere Kanten (in Z-Höhe slab_z) eines Slabs
    # mit seen werden doppelte Kanten vermieden. Diese entstehen dadurch, 
    # dass kanten zu zwei verschiedenen flächen gehören und unterschiedliche richtungen haben und somit doppelt ausgewertet werden
    seen = set()
    top_edges = []
    edge_explorer = TopExp_Explorer()
    edge_explorer.Init(slab_shape, TopAbs_EDGE)

    while edge_explorer.More():
        edge = edge_explorer.Current()
        adaptor = BRepAdaptor_Curve(edge)
        start_of_edge = adaptor.FirstParameter()
        end_of_edge = adaptor.LastParameter()
        mid_of_edge = 0.5 * (start_of_edge + end_of_edge)
        mid_point = adaptor.Value(mid_of_edge)

        if abs(mid_point.Z() - slab_z) <= Z_TOLERANCE:
            start_point = adaptor.Value(start_of_edge)
            end_point = adaptor.Value(end_of_edge)
            coords_start_point = (start_point.X(), start_point.Y(), start_point.Z())
            coords_end_point = (end_point.X(), end_point.Y(), end_point.Z())
            key = tuple(sorted((coords_start_point, coords_end_point)))

            if key not in seen:
                seen.add(key)
                top_edges.append((edge, mid_point))

        edge_explorer.Next()

    return top_edges



def compute_plate_for_edge(edge, mid, slab_shape, wall_shapes):
    # Erstellt Safe- und Unsafe-Plates für die gesammelten Kanten in get_top_edges und liefert Endpunkte von diesen

    # Richtung in XY-Ebene
    curve, start_of_edge, end_of_edge = BRep_Tool.Curve(edge)
    start_point = curve.Value(start_of_edge)
    end_point = curve.Value(end_of_edge)
    vector_tangential = gp_Vec(end_point.X()-start_point.X(), end_point.Y()-start_point.Y(), 0.0)
    vector_tangential.Normalize()
    vector_rechtwinklig = gp_Vec(-vector_tangential.Y(), vector_tangential.X(), 0.0)
    vector_rechtwinklig.Normalize()
    
    # Testpunkt verschieben und prüfen, ob in die richtige Richtung verschoben wurde
    point_test = gp_Pnt(mid.X()+vector_rechtwinklig.X()*OFFSET, mid.Y()+vector_rechtwinklig.Y()*OFFSET, mid.Z())
    vertex = BRepBuilderAPI_MakeVertex(point_test).Shape()
    distance = BRepExtrema_DistShapeShape(vertex, slab_shape).Value()
    direction = vector_rechtwinklig if distance <= DIST_THRESHOLD else vector_rechtwinklig.Reversed()
    shift = direction.Multiplied(OFFSET)
    
    # Kante verschieben um OFFSET und in der Länge trimmen um LENGTH_DIFF
    translation = gp_Trsf()
    translation.SetTranslation(shift)
    moved = BRepBuilderAPI_Transform(edge, translation, True).Shape()
    curve_t, tu1, tu2 = BRep_Tool.Curve(moved) #or BRep_Tool.Curve(edge)
    q1 = curve_t.Value(tu1)
    q2 = curve_t.Value(tu2)
    dir_vec = gp_Vec(q1, q2)
    dir_vec.Normalize()
    p1n = gp_Pnt(q1.X()+dir_vec.X()*LENGTH_DIFF, q1.Y()+dir_vec.Y()*LENGTH_DIFF, q1.Z()+dir_vec.Z()*LENGTH_DIFF)
    p2n = gp_Pnt(q2.X()-dir_vec.X()*LENGTH_DIFF, q2.Y()-dir_vec.Y()*LENGTH_DIFF, q2.Z()-dir_vec.Z()*LENGTH_DIFF)
    trimmed = BRepBuilderAPI_MakeEdge(p1n, p2n).Edge()
    
    # Plate extrudieren und cutten
    plate = BRepPrimAPI_MakePrism(trimmed, gp_Vec(0,0,MIN_WALL_HEIGHT), True).Shape()

    # Sonderfall: Wenn keine Wände vorhanden sind, wird die gesamte Plate als unsafe betrachtet
    if not wall_shapes:
        unsafe = plate
        safe = BRepBuilderAPI_MakeVertex(mid).Shape()
        coords = (p1n.X(), p1n.Y(), p1n.Z(), p2n.X(), p2n.Y(), p2n.Z())
        unsafe_list = [coords]
        print(f"berechnete Koordinaten: ({coords[0]:.3f}, {coords[1]:.3f}, {coords[2]:.3f}, "f"{coords[3]:.3f}, {coords[4]:.3f}, {coords[5]:.3f})")
        return safe, unsafe, plate, unsafe_list
    

    unsafe = plate
    
    for wall in wall_shapes:
        try:
            # Prüfe zunächst, ob die Shapes gültige Geometrie haben
            if shape_is_empty(unsafe) or shape_is_empty(wall):
                continue

            # Versuch unsafe (Plate) mit der Wand zu schneiden
            cutter = BRepAlgoAPI_Cut(unsafe, wall)
            if not cutter.IsDone():
                print(f"Cut nicht abgeschlossen für Wand {wall}, überspringe")
                continue
            new_unsafe = cutter.Shape()
            if new_unsafe is None:
                print(f"Kein Shape von cutter bei Wand {wall} zurückgegeben, überspringe")
                continue
            unsafe = new_unsafe

            # Erzeuge sichere Variante: plate minus updated unsafe
            safe_cut = BRepAlgoAPI_Cut(plate, unsafe)
            if not safe_cut.IsDone():
                print(f"Cut nicht abgeschlossen für plate und unsafe, überspringe safe-Erstellung")
                continue
            safe_cut.Build()
            safe_shape = safe_cut.Shape()
            if safe_shape is None:
                print(f"Kein Shape von safe_cut zurückgegeben, überspringe")
                continue
            safe = safe_shape

        except Exception as e:
            print(f"Fehler im Cut-Prozess mit Wand {wall}: {e}, überspringe")
            continue


    
    # Unsichere Teilabschnitte ermitteln: Extrahiere horizontale Kanten des Unsafe-Teils
    unsafe_list = []
    exp_edge = TopExp_Explorer()
    exp_edge.Init(unsafe, TopAbs_EDGE)
    seen = set()
    while exp_edge.More():
        edge = exp_edge.Current()
        adaptor = BRepAdaptor_Curve(edge)
        anfang_kante = adaptor.FirstParameter()
        ende_kante = adaptor.LastParameter()
        mitte_kante = 0.5 * (anfang_kante + ende_kante)
        punkt_mitte_kante = adaptor.Value(mitte_kante)
        # horizontale Kanten in Projektionshöhe des Slabs
        if abs(punkt_mitte_kante.Z() - mid.Z()) <= Z_TOLERANCE:
            punkt_anfang_kante = adaptor.Value(anfang_kante)
            punkt_ende_kante = adaptor.Value(ende_kante)
            coords = (punkt_anfang_kante.X(), punkt_anfang_kante.Y(), punkt_anfang_kante.Z(), punkt_ende_kante.X(), punkt_ende_kante.Y(), punkt_ende_kante.Z())
            key = tuple(sorted(((coords[0],coords[1]),(coords[3],coords[4]))))
            if key not in seen:
                seen.add(key)
                unsafe_list.append(tuple(round(v,3) for v in coords))
        exp_edge.Next()

    # Bei gemischtem Zustand (unsafe und safe nebeneinander) wird eine Schnittkante zwischen den beiden gesucht
    if not shape_is_empty(unsafe) and not shape_is_empty(safe):
        # Schnittkante zwischen Unsafe und Safe
        section = BRepAlgoAPI_Section(unsafe, safe)
        section.Approximation(True)
        section.ComputePCurveOn1(True)
        section.Build()
        if section.IsDone():
            verts = []
            exp_v = TopExp_Explorer()
            exp_v.Init(section.Shape(), TopAbs_VERTEX)
            while exp_v.More():
                p = BRep_Tool.Pnt(exp_v.Current())
                verts.append((p.X(), p.Y(), p.Z()))
                exp_v.Next()
            # Wähle für jeden X,Y den niedrigsten Z-Wert als Übergang
            minz = {}
            for x, y, z in verts:
                key_xy = (round(x,3), round(y,3))
                if key_xy in minz:
                    minz[key_xy] = min(minz[key_xy], z)
                else:
                    minz[key_xy] = z
            for (x, y), zmin in minz.items():
                key = ((x, y), (x, y))
                if key not in seen:
                    seen.add(key)
                    # Füge den Punkt als Start- und Endpunkt gleichen Werts hinzu
                    unsafe_list.append((x, y, zmin, x, y, zmin))

    # Ausgabe der gefundenen Endpunkte
    for x1,y1,z1,x2,y2,z2 in unsafe_list:
        print(f"berechnete Koordinaten: ({x1:.3f}, {y1:.3f}, {z1:.3f}, {x2:.3f}, {y2:.3f}, {z2:.3f})")
    
    return safe, unsafe, plate, unsafe_list



def visualize_plates(display, safe, unsafe, col_green, col_red):
    #Zeigt Safe- und Unsafe-Plates in der Anzeige
    for plate, color in [(safe, col_green), (unsafe, col_red)]:
        ais = AIS_Shape(plate)
        ais.SetColor(color)
        display.Context.Display(ais, False)


def write_csv(filename, coords_list):
    #Speichert Koordinaten und Typ in eine CSV
    with open(filename, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['x_start','y_start','z_start','x_end','y_end','z_end','Typ'])
        print()
        print("Diese Koordinaten wurden in die csv-Datei exportiert:")
        for coords in coords_list:
            z = coords[2]
            if z >= GRENZE_GELAENDER:
                typ = 'Geruest' if z >= GRENZE_GERUEST else 'Gelaender'
                rounded_coords = [f"{c:.3f}" for c in coords]
                w.writerow([*rounded_coords, typ])
                print(*rounded_coords, typ)


def visualize_unsafe_coords(unsafe_coords, display, col_magenta):
    #Visualisiert die unsicheren Kanten-Endpunkte als Kugeln
    for coords in unsafe_coords:
        x1, y1, z1, x2, y2, z2 = coords
        for x, y, z in [(x1, y1, z1), (x2, y2, z2)]:
            pnt = gp_Pnt(x, y, z)
            if z >= GRENZE_GELAENDER:
                kugel = BRepPrimAPI_MakeSphere(pnt, RADIUS).Shape()
                display.DisplayShape(kugel, update=False, color=col_magenta)




# --- Hauptfunktion ---
def finde_absturzkanten(ifc_path: str, settings):
    directory, base, slab_shapes, wall_shapes = load_ifc_data(ifc_path, settings)
    display, start_display, col_red, col_green, col_slab, col_wall, col_magenta = init_visualization()
    unsafe_coords = []

    for wall in wall_shapes:
        # Wand transparent anzeigen
        ais = AIS_Shape(wall)
        ais.SetColor(col_wall)
        ais.SetTransparency(0.7)  # 0.0 = undurchsichtig, 1.0 = komplett transparent
        display.Context.Display(ais, True)

    for slab in slab_shapes:
        # Slab anzeigen
        ais = AIS_Shape(slab)
        ais.SetColor(col_slab)
        display.Context.Display(ais, True)

        _,_,_,_,_,zmax = bbox_bounds(slab)
        
        for edge, mid in get_top_edges(slab, zmax):
            safe, unsafe, plate, coords_list = compute_plate_for_edge(edge, mid, slab, wall_shapes)
            visualize_plates(display, safe, unsafe, col_green, col_red)
            unsafe_coords.extend(coords_list)

        visualize_unsafe_coords(unsafe_coords, display, col_magenta)
   
    csv_file = os.path.join(directory, f"Punkte_{base}.csv")
    write_csv(csv_file, unsafe_coords)
    print()
    print("Speicherpfad der csv-Datei:")
    print(csv_file)

    display.FitAll()
    display.View.SetZoom(1.0)
    start_display()

    return unsafe_coords


class TextRedirector(object):
    #Hilfsklasse, um print-Ausgaben in Tkinter Fenster umzuleiten
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.configure(state='normal')
        self.widget.insert('end', s)
        self.widget.see('end')
        self.widget.configure(state='disabled')

    def flush(self):
        pass

def run_finde_absturzkanten(ifc_file, settings, button):
    button.config(state='disabled')
    try:
        finde_absturzkanten(ifc_file, settings)
    except Exception as e:
        print(f"Fehler: {e}")
    button.config(state='normal')


# Hilfsfunktion für Ressourcenpfad (für Icon u.a.)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)


def start_gui():
    root = tk.Tk()
    root.title("Absturzkanten Finder")
    icon = resource_path('Suva_RGB_orange.ico')
    try:
        root.iconbitmap(icon)
    except tk.TclError:
        pass

    # DPI Awareness für hochauflösende Displays (Windows)
    try:
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(size=16)
    text_font = tkfont.Font(family="Consolas", size=16)

    # anpassen der GUI-Grösse wenn Fenster verändert wird
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    frame = tk.Frame(root)
    frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    frame.rowconfigure(2, weight=1)
    frame.columnconfigure(0, weight=1)

    # Eingabefeld für HEIGHT_SITE
    lbl_height = tk.Label(frame, text="Höhe Gelände [m]:", font=default_font)
    lbl_height.grid(row=0, column=0, sticky="w", padx=(0,10))  # Abstand von 10 Pixeln nach rechts
    var_height = tk.StringVar(value=str(HEIGHT_SITE))
    entry_height = tk.Entry(frame, textvariable=var_height, font=default_font, width=10)
    entry_height.grid(row=0, column=0, sticky="w", padx=(200,10))

    btn_select = tk.Button(frame, text="IFC Datei wählen", font=default_font)
    btn_select.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10,0))

    txt_output = tk.Text(frame, height=18, width=80, state='disabled', wrap='word', font=text_font)
    txt_output.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10,0))

    # Scrollbar hinzufügen
    scrollbar = tk.Scrollbar(frame, command=txt_output.yview)
    scrollbar.grid(row=2, column=2, sticky='ns', pady=(10,0))
    txt_output['yscrollcommand'] = scrollbar.set

    sys.stdout = TextRedirector(txt_output)
    sys.stderr = TextRedirector(txt_output)

    def on_select():
        global HEIGHT_SITE, GRENZE_GELAENDER, GRENZE_GERUEST
        try:
            HEIGHT_SITE = float(var_height.get())
        except Exception:
            print("Ungültiger Wert für die Geländehöhe, bitte gültige Zahl eingeben.")
            return
        # Grenzen für Geländer und Gerüst basierend auf HEIGHT_SITE durch Eingabe neu berechnen
        GRENZE_GELAENDER = 2.0 + HEIGHT_SITE
        GRENZE_GERUEST = 3.0 + HEIGHT_SITE

        ifc_file = filedialog.askopenfilename(
            title="IFC Datei waehlen",
            filetypes=[("IFC Dateien", "*.ifc"), ("Alle Dateien", "*")]
        )
        if not ifc_file:
            print("Keine IFC Datei gewaehlt.")
            return
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_PYTHON_OPENCASCADE, True)
        # In Thread ausführen, damit UI nicht blockiert
        threading.Thread(target=run_finde_absturzkanten, args=(ifc_file, settings, btn_select), daemon=True).start()

    btn_select.config(command=on_select)

    root.mainloop()

if __name__ == "__main__":
    start_gui()

# Code als .exe speichern:
# 1. anaconda prompt öffnen
# 2. cd c:\Users\wpx619\.AA_IP4-SUVA_VSCode
# 3. conda activate c:\Users\wpx619\.AA_IP4-SUVA_VSCode\.conda
# 4. (wenn nicht bereits installiert) pip install pyinstaller 
# 5. pyinstaller --windowed --icon=Suva_RGB_orange.ico --add-data "Suva_RGB_orange.ico;." absturzkanten_04_mp.py