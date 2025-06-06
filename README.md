# Absturzkanten-Finder

Dieses Python-Tool analysiert IFC-Dateien und erkennt potenziell unsichere Absturzkanten basierend auf der Geometrie von Bodenplatten (Slabs) und Wänden (Walls). Die resultierenden unsicheren Kanten werden visuell dargestellt und als CSV-Datei exportiert und je nach Distanz zum Gelände (IfcSite) klassifiziert nach Geländer oder Gerüst.

## Funktionen

- Einlesen und Verarbeitung von IFC-Dateien mit `ifcopenshell`
- 3D-Geometrieverarbeitung mit PythonOCC
- Visualisierung der Geometrien und Ergebnisse (inkl. farblicher Unterscheidung von sicheren/unsicheren Bereichen)
- Benutzeroberfläche mit Tkinter
- Export der gefundenen Kantenpunkte als CSV mit Klassifikation
- Konfigurierbare Geländehöhe zur automatischen Unterscheidung von Gerüst und Geländer

## Voraussetzungen

- Python 3.8 oder höher (getestet mit Python 3.12.9)
- Empfohlen: Nutzung innerhalb einer Conda-Umgebung

### Benötigte Python-Pakete

- `ifcopenshell` (https://anaconda.org/conda-forge/ifcopenshell)
- `pythonocc-core` (https://anaconda.org/conda-forge/pythonocc-core)
- `pyinstaller` (nur für die Erstellung einer .exe)

### Installation mit Conda (empfohlen)

Erstelle und aktiviere eine neue Umgebung:

```bash
conda create -n absturzkanten python=3.8
conda activate absturzkanten

